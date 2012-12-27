<%@ Page Language="vb" AutoEventWireup="false" Codebehind="DBClient.aspx.vb" Inherits="DBManage.DBClient" ValidateRequest="false"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>Database Client</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<LINK href="my.css" type="text/css" rel="stylesheet">
		<script language="javascript">
		function ShowSql()
		{
			document.getElementById("txtQuery").value="Select * from "+document.getElementById("DrpTables").value;
		}
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<asp:TextBox id="txtQuery" style="Z-INDEX: 100; LEFT: 32px; POSITION: absolute; TOP: 152px" runat="server"
				Height="44px" Width="666px"></asp:TextBox>
			<asp:Label id="LblTime" style="Z-INDEX: 114; LEFT: 376px; POSITION: absolute; TOP: 24px" runat="server"
				Width="312px" Height="20px" BorderWidth="1px"></asp:Label>
			<asp:Label id="Label1" style="Z-INDEX: 113; LEFT: 264px; POSITION: absolute; TOP: 24px" runat="server"
				Width="48px" Height="24px">Time</asp:Label>
			<TABLE id="Table1" style="Z-INDEX: 112; LEFT: 32px; POSITION: absolute; TOP: 248px; HEIGHT: 43px"
				cellSpacing="1" cellPadding="1" width="745" border="0">
				<TR>
					<TD style="HEIGHT: 20px"></TD>
				</TR>
				<TR>
					<TD vAlign="top" align="left">
						<asp:DataGrid id="DataGrid1" runat="server" Width="740px" Height="50px" GridLines="Vertical" CellPadding="3"
							BorderWidth="1px" BorderColor="#999999" BorderStyle="Solid" ForeColor="Black" BackColor="White"
							PageSize="30" AllowPaging="True">
							<SelectedItemStyle Font-Bold="True" ForeColor="White" BackColor="#000099"></SelectedItemStyle>
							<AlternatingItemStyle BackColor="#CCCCCC"></AlternatingItemStyle>
							<ItemStyle Font-Size="XX-Small" Font-Names="Verdana" HorizontalAlign="Left"></ItemStyle>
							<HeaderStyle Font-Size="Smaller" Font-Names="Verdana" Font-Bold="True" HorizontalAlign="Center"
								ForeColor="White" BackColor="Black"></HeaderStyle>
							<FooterStyle BackColor="#CCCCCC"></FooterStyle>
							<PagerStyle HorizontalAlign="Center" ForeColor="Black" Position="Top" BackColor="#999999" Mode="NumericPages"></PagerStyle>
						</asp:DataGrid></TD>
				</TR>
			</TABLE>
			<asp:Button id="BtnRun" style="Z-INDEX: 101; LEFT: 704px; POSITION: absolute; TOP: 168px" runat="server"
				Height="24px" Width="64px" Text="Run" CssClass="CmdButton"></asp:Button>
			<asp:TextBox id="txtDDL" style="Z-INDEX: 102; LEFT: 32px; POSITION: absolute; TOP: 72px" runat="server"
				Height="56px" Width="664px" TextMode="MultiLine"></asp:TextBox>
			<asp:Label id="Label2" style="Z-INDEX: 103; LEFT: 32px; POSITION: absolute; TOP: 136px" runat="server"
				Height="16px" Width="112px">Select Query</asp:Label>
			<asp:Label id="Label3" style="Z-INDEX: 104; LEFT: 32px; POSITION: absolute; TOP: 56px" runat="server"
				Height="16px" Width="208px">DDL Statements</asp:Label>
			<asp:Button id="BtnExecute" style="Z-INDEX: 105; LEFT: 704px; POSITION: absolute; TOP: 104px"
				runat="server" Height="24px" Width="66px" Text="Execute" CssClass="CmdButton"></asp:Button>
			<asp:Label id="LblError" style="Z-INDEX: 106; LEFT: 32px; POSITION: absolute; TOP: 200px" runat="server"
				Height="40px" Width="664px" BorderColor="#C00000" BorderWidth="1px"></asp:Label>
			<asp:Label id="lblTable" style="Z-INDEX: 107; LEFT: 32px; POSITION: absolute; TOP: 8px" runat="server"
				Width="48px" Height="24px">Tables</asp:Label>
			<asp:DropDownList id="DrpTables" style="Z-INDEX: 108; LEFT: 32px; POSITION: absolute; TOP: 32px" runat="server"
				Width="224px" Height="24px"></asp:DropDownList>
			<asp:Label id="LblTableCount" style="Z-INDEX: 110; LEFT: 376px; POSITION: absolute; TOP: 48px"
				runat="server" Width="312px" Height="20px" BorderWidth="1px"></asp:Label>
			<asp:Button id="BtnStructure" style="Z-INDEX: 111; LEFT: 264px; POSITION: absolute; TOP: 48px"
				runat="server" Width="96px" Text="Show Structure" CssClass="CmdButton"></asp:Button>
		</form>
	</body>
</HTML>
