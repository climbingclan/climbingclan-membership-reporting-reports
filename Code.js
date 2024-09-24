const scriptProperties = PropertiesService.getScriptProperties();

const server = scriptProperties.getProperty('cred_server');
const port = parseInt(scriptProperties.getProperty('cred_port'), 10);
const dbName = scriptProperties.getProperty('cred_dbName');
const username = scriptProperties.getProperty('cred_username');
const password = scriptProperties.getProperty('cred_password');
const url = `jdbc:mysql://${server}:${port}/${dbName}`;
const apidomain = scriptProperties.getProperty('cred_apidomain');
const apiusername = scriptProperties.getProperty('cred_apiusername');
const apipassword = scriptProperties.getProperty('cred_apipassword');

function readData() {
 var conn = Jdbc.getConnection(url, username, password);
 var stmt = conn.createStatement();


 var spreadsheet = SpreadsheetApp.getActive();
 var sheet = spreadsheet.getSheetByName('Dashboard');
 //var cell = sheet.getRange('B4').getValues();


 var results = stmt.executeQuery('select distinct  `first_name` "Forenames",`last_name` "Surname", `admin-dob` "Dob", `shipping_address_1` "Address 1",`shipping_address_2` "Address 2",`shipping_city` "Town", `shipping_postcode` "Postcode",`billing_email` "Email",`admin-phone-number` "Home Tel.",pd.user_id "Club Ref", `admin-bmc-membership-number`  "BMC Ref",`admin-membership-type` AS "Membership Type", (SELECT IFNULL( (select `dbi`.`committee_current` from `wp_member_db` dbi where dbi.id = `db`.`id` LIMIT 1) ,"")) AS `Membership Title`, (SELECT (CASE WHEN `dbi`.`committee_current`="secretary" THEN "Main Club Contact" WHEN `dbi`.`committee_current`="chair" THEN "Communications Contact" ELSE "Club Member" END) FROM `wp_member_db` dbi where dbi.id = `db`.`id` LIMIT 1) AS `Contact Type`, pd.order_id "Order Ref" from wp_member_db db JOIN wp_order_product_customer_lookup pd on pd.user_id = db.id where product_id=2118 AND db.cc_member="yes" AND status in ("wc-processing", "wc-completed") AND YEAR(pd.order_created)>YEAR(CURDATE())-1 order by `first_name` ASC;');
  //console.log(results);
 var metaData=results.getMetaData();
  var numCols = metaData.getColumnCount();
 var spreadsheet = SpreadsheetApp.getActive();
 var sheet = spreadsheet.getSheetByName('BMC-MSO-Report');
 sheet.clearContents();
  //sheet.appendRow(["Processed Members"]);
 var arr=[];
  for (var col = 0; col < numCols; col++) {
   arr.push(metaData.getColumnLabel(col + 1));
 }
 // https://stackoverflow.com/questions/10585029/parse-an-html-string-with-js
  sheet.appendRow(arr);
 while (results.next()) {
 arr=[];
 for (var col = 0; col < numCols; col++) {
   arr.push(results.getString(col + 1));
 }
 sheet.appendRow(arr);
 }
sheet.autoResizeColumns(1, numCols+1);

// Postcode

 var spreadsheet = SpreadsheetApp.getActive();
 var sheet = spreadsheet.getSheetByName('Dashboard');
 //var cell = sheet.getRange('B4').getValues();


 var results = stmt.executeQuery('select distinct  `first_name` "Forenames",`last_name` "Surname", `shipping_postcode` "Postcode",`admin-membership-type` AS "Membership Type" FROM wp_member_db db JOIN wp_order_product_customer_lookup pd on pd.user_id = db.id where `shipping_postcode` IS NOT NULL order by `first_name` ASC;');
  //console.log(results);
 var metaData=results.getMetaData();
  var numCols = metaData.getColumnCount();
 var spreadsheet = SpreadsheetApp.getActive();
 var sheet = spreadsheet.getSheetByName('IMD-Report');
 sheet.clearContents();
  //sheet.appendRow(["Processed Members"]);
 var arr=[];
  for (var col = 0; col < numCols; col++) {
   arr.push(metaData.getColumnName(col + 1));
 }
 // https://stackoverflow.com/questions/10585029/parse-an-html-string-with-js
  sheet.appendRow(arr);
 while (results.next()) {
 arr=[];
 for (var col = 0; col < numCols; col++) {
   arr.push(results.getString(col + 1));
 }
 sheet.appendRow(arr);
 }
sheet.autoResizeColumns(1, numCols+1);


// to_process Report
//
 var spreadsheet = SpreadsheetApp.getActive();
 var sheet = spreadsheet.getSheetByName('To Process');
 sheet.clearContents();
 sheet.appendRow(["Members to process"]);


// start of to_process function
function to_process(querystring, title)
{
sheet.appendRow([title]);
//sheet.getLastRow(row+1,1).setFontWeight("bold");
//sheet.getRange(row+1,1).setFontWeight("bold");
//var lastRow = sheet.getLastRow(row+1,1);

 var results = stmt.executeQuery('select distinct  `first_name` "Forenames",`last_name` "Surname", `shipping_address_1` "Address 1",`shipping_address_2` "Address 2",`shipping_city` "Town", `shipping_postcode` "Postcode",`billing_email` "Email",`admin-phone-number` "Home Tel.",pd.order_item_name "Membership Name", pd.variation_id "Variation ID", user_id "User ID", pd.order_id "Order ID", (SELECT CONCAT("https://www.climbingclan.com/wp-admin/post.php?post=",pd.order_id,"&action=edit") ) as "Order Edit" from wp_member_db db JOIN wp_order_product_customer_lookup pd on pd.user_id = db.id where product_id=2118 AND status="wc-processing" AND YEAR(pd.order_created)>YEAR(CURDATE())-1  AND `variation_id` LIKE "%'+ querystring + '%" order by `first_name` ASC');
  //console.log(results);
 var metaData=results.getMetaData();
  var numCols = metaData.getColumnCount();
 var arr=[];
  for (var col = 0; col < numCols; col++) {
   arr.push(metaData.getColumnName(col + 1));
 }
 // https://stackoverflow.com/questions/10585029/parse-an-html-string-with-js
  sheet.appendRow(arr);
 while (results.next()) {
 arr=[];
 for (var col = 0; col < numCols; col++) {
   arr.push(results.getString(col + 1));
 }
 sheet.appendRow(arr);
 }
sheet.autoResizeColumns(1, numCols+1);



} //end of to_process

// full options
//help online beforehand,help at sign-in,help around announcements and cake time,do announcements,help online afterwards,be event director for the evening

to_process("2360", "Benefactor Members");
to_process("2361", "Members with Mugs");
to_process("2357", "Regular Members");
//volunteering("online afterwards", "Online Afterwards");
//volunteering("do announcements", "Do Announcements");
//volunteering("be event dir", "Be Event Director");

//var range = SpreadsheetApp.getActive().getRange("Volunteering!C2:C150");
//range.insertCheckboxes();
//sheet.autoResizeColumns(1, numCols+1);





sheet.autoResizeColumns(1, numCols+1);

// Processed Report
//
 var spreadsheet = SpreadsheetApp.getActive();
 var sheet = spreadsheet.getSheetByName('Processed');
 sheet.clearContents();
 sheet.appendRow(["Processed Members"]);


// start of Processed function
function processed(querystring, title)
{
sheet.appendRow([title]);
//sheet.getLastRow(row+1,1).setFontWeight("bold");
//sheet.getRange(row+1,1).setFontWeight("bold");
//var lastRow = sheet.getLastRow(row+1,1);

 var results = stmt.executeQuery('select distinct  `first_name` "Forenames",`last_name` "Surname", pd.order_item_name "Membership Name", pd.variation_id "Variation ID", user_id "User ID", pd.order_id "Order ID" from wp_member_db db JOIN wp_order_product_customer_lookup pd on pd.user_id = db.id where product_id=2118 AND status="wc-completed" AND YEAR(pd.order_created)>YEAR(CURDATE())-1 AND `variation_id` LIKE "%'+ querystring + '%" order by `first_name` ASC');
  //console.log(results);
 var metaData=results.getMetaData();
  var numCols = metaData.getColumnCount();
 var arr=[];
  for (var col = 0; col < numCols; col++) {
   arr.push(metaData.getColumnName(col + 1));
 }
 // https://stackoverflow.com/questions/10585029/parse-an-html-string-with-js
  sheet.appendRow(arr);
 while (results.next()) {
 arr=[];
 for (var col = 0; col < numCols; col++) {
   arr.push(results.getString(col + 1));
 }
 sheet.appendRow(arr);
 }



} //end of processed

// full options
//help online beforehand,help at sign-in,help around announcements and cake time,do announcements,help online afterwards,be event director for the evening

processed("2360", "Benefactor Members");
processed("2361", "Members with Mugs");
processed("2357", "Regular Members");
//volunteering("online afterwards", "Online Afterwards");
//volunteering("do announcements", "Do Announcements");
//volunteering("be event dir", "Be Event Director");

//var range = SpreadsheetApp.getActive().getRange("Volunteering!C2:C150");
//range.insertCheckboxes();
//sheet.autoResizeColumns(1, numCols+1);





sheet.autoResizeColumns(1, numCols+1);

//close SQL
results.close();
stmt.close();



//end read data function
}


//ScriptApp.newTrigger('readData')
//.timeBased()
//.everyMinutes(30)
//.create();
