# Query all in porcess contract by end date
import mysql.connector
import pandas as pd
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
import json
from google.cloud import bigquery
import os

# Connect to the Tencent Cloud MySQL database
connection = mysql.connector.connect(
    host='sg-cdb-p3cg8xwh.sql.tencentcdb.com',
    user='ispl_readonly',
    password='Lylo12345678##',
    port = 27085,
    database='ispl_prod'
)

# Query the data
query = '''
with report_days as (
	with recursive date_table as(
		SELECT
		date_sub(curdate(), interval 1 day) as report_day
		union ALL
		select date_add(report_day, interval 1 day) report_day
		from date_table
		where report_day < date_add(curdate(), interval 2 year)
		)
	select report_day
	from date_table
),
all_active_contract as (
	SELECT 
		CASE
	    	WHEN ct.cont_stage = "Active" and cont_type = 'newcontract' then 'newcontract'
	    	WHEN ct.cont_stage = "Active" and cont_type = 'recontract' then 'recontract'
	    	when ct.cont_stage = "ced" then 'ced'
	    	when ct.cont_stage = "Terminate" then 'terminate'
	    	when ct.cont_stage = "tempreturned" then 'tempreturned'
	    	Else 'others'
		 End as Type,
	     CASE
	        WHEN ct.cont_name LIKE 'C%' THEN 'Lumens'
	        WHEN ct.cont_name LIKE 'FR%' THEN 'Focus'
	        ELSE 'Others'
	    END AS 'BU',
	    date(date_add(ct.cont_CreatedDate, interval 8 hour)) AS 'Reportday',
	--  ct.cont_CreatedDate AS 'Created Time',
	    c.Comp_Name AS 'Driver Name',
	    v.vehi_Make AS 'Model',
	    v.vehi_Model AS 'Model Type',
	    ct.cont_name AS 'Contract Number',
	    ct.cont_subtype,ct.cont_contstartdate AS 'Start Date',
	    ct.cont_contenddate AS 'Contract End Date',
	    v.vehi_PlateNo AS 'Car Plate',
		case 
			when u.User_Logon is null  then uu.User_Logon
			else u.User_Logon
		end AS "RM",
	    uu.User_Logon AS "Sales"
	From crm_contract ct 
	LEFT JOIN crm_company c ON c.Comp_CompanyId = ct.cont_companyid 
	LEFT JOIN crm_vehicle v on ct.cont_vehicleid = v.vehi_VehicleID
	LEFT JOIN crm_contract_detail cc ON cc.code_contractid = ct.cont_contractid
	LEFT JOIN crm_user u on ct.cont_rmincharge = u.User_UserId
	LEFT JOIN crm_user uu on ct.cont_saleincharge = uu.User_UserId
	WHERE 
	--  ct.cont_CheckedOut = 'Y' AND
		lower(CONT_TYPE) != 'interbu' AND
		ct.cont_stage = "Active" AND
		ct.cont_Deleted IS NULL AND
		cc.code_itemtype = 'rental' AND
	--	ct.cont_stage = "Active" AND
		DATE(DATE_ADD(ct.cont_contenddate, INTERVAL 8 HOUR)) 
		BETWEEN CURDATE() AND DATE_ADD(CURDATE(), INTERVAL 2 YEAR)
),
contract_by_day as (
	select 
	date(`Contract End Date`) `Contract End Date`,
	BU,
	`Driver Name`,
	`Model`,
	`Model Type`,
	date(`Start Date`) `Start Date`,
	GREATEST(TIMESTAMPDIFF(MONTH, `Start Date`, `Contract End Date`), 1) AS tenure,
	 `Car Plate`,
	RM, Sales,`Contract Number`,
	Type
	from all_active_contract
	where Type != 'others'
	order by `Contract End Date`,BU
)
select * from contract_by_day
'''

print ("Executing SQL Query...")
data_frame = pd.read_sql(query, connection)
print (data_frame.head())

# Close the connection
connection.close()

# 改成输出 xlsx 文件，并指定 sheet 名字
output_path = "contract_by_enddate.xlsx"
data_frame.to_excel(output_path, sheet_name="Detail", index=False)
print(f"Excel file saved to {output_path}")

# SharePoint 上传部分不变，只是换了文件扩展名
site_url = "https://lumensautopl.sharepoint.com/sites/Contract_by_enddate/"
username = "bi-team@lumens.sg"
password = "Lumens@2022"


target_folder_url = "Shared%20Documents/contact_by_enddate/"

ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))
try:
    with open(output_path, "rb") as file_content:
        target_folder = ctx.web.get_folder_by_server_relative_url(target_folder_url)
        target_folder.upload_file("contract_by_enddate.xlsx", file_content.read()).execute_query()
    print(f"\033[32m[Ok] File uploaded to SharePoint successfully! {site_url + target_folder_url}\033[0m")
except json.JSONDecodeError as e:
    print(f"JSON decode error: {e}")
except Exception as ex:
    print(f"An error occurred: {ex}")
