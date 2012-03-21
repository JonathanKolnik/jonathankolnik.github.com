<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>Untitled Document</title>
</head>

<body><?php
$dbhost = 'localhost';
$dbuser = 'toy1014610254501';
$dbpass = 'jk5t3En7';
$conn = mysql_connect($dbhost, $dbuser, $dbpass) or die('Error connecting to mysql');
$dbname = 'databasename';
mysql_select_db($dbname);
$query  = "UPDATE wp_options SET option_value = 'http://www.toyboxapparel.com' WHERE `wp_options`.`option_id` =1 AND `wp_options`.`blog_id` =0 AND CONVERT( `wp_options`.`option_name` USING utf8 ) = 'siteurl' LIMIT 1 ;";
$result = mysql_query($query);
echo $result;
$query = "UPDATE wp_options SET option_value = 'http://www.toyboxapparel.com' WHERE `wp_options`.`option_id` =39 AND `wp_options`.`blog_id` =0 AND CONVERT( `wp_options`.`option_name` USING utf8 ) = 'home' LIMIT 1 ;";
$result = mysql_query($query);
echo "Done.";
mysql_close($conn);
?>
</body>
</html>