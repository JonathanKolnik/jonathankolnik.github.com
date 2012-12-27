<html>
<head>
<title>Castles Unlimited Meet Our Agents</title>
<link href="template_four.css" type="text/css" rel="stylesheet">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript1.2" src="functions.js"></script>
<style type="text/css">
<!--
.style1 {color: #980108}
-->
</style>
</head>

<table width='100%' border=0>
<tr><td>
<?php
if( isset($_REQUEST['agent_id']))
{
$agent_id=$_REQUEST['agent_id'];
include("http://10.1.1.70/agent_roster/index.php?agent_id=$agent_id");
}
else
{
include("http://10.1.1.70/agent_roster/index.php?company_id=862");
}
?>
</td></tr>

</table>

</body>
</html>
