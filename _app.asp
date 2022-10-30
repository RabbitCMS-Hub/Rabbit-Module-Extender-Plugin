<%
'**********************************************
'**********************************************
'               _ _                 
'      /\      | (_)                
'     /  \   __| |_  __ _ _ __  ___ 
'    / /\ \ / _` | |/ _` | '_ \/ __|
'   / ____ \ (_| | | (_| | | | \__ \
'  /_/    \_\__,_| |\__,_|_| |_|___/
'               _/ | Digital Agency
'              |__/ 
' 
'* Project  : RabbitCMS
'* Developer: <Anthony Burak DURSUN>
'* E-Mail   : badursun@adjans.com.tr
'* Corp     : https://adjans.com.tr
'**********************************************
' LAST UPDATE: 28.10.2022 15:33 @badursun
'**********************************************

Class Rabbit_Module_Extender_Plugin
	Private PLUGIN_CODE, PLUGIN_DB_NAME, PLUGIN_NAME, PLUGIN_VERSION, PLUGIN_CREDITS, PLUGIN_GIT, PLUGIN_DEV_URL, PLUGIN_FILES_ROOT, PLUGIN_ICON, PLUGIN_REMOVABLE, PLUGIN_ROOT, PLUGIN_FOLDER_NAME
	Private MODULE_ROOT, LOADED_MODULES, MODULE_STATUS
	Private ORDER_JS, ORDER_CSS

	'---------------------------------------------------------------
	' Register Class
	'---------------------------------------------------------------
	Public Property Get class_register()
		DebugTimer ""& PLUGIN_CODE &" class_register() Start"
		
		' Check Register
		'------------------------------
		If CheckSettings("PLUGIN:"& PLUGIN_CODE &"") = True Then 
			DebugTimer ""& PLUGIN_CODE &" class_registered"
			Exit Property
		End If

		' Check And Create Table
		'------------------------------
		Dim PluginTableName
			PluginTableName = "tbl_plugin_" & PLUGIN_DB_NAME
    	
    	If TableExist(PluginTableName) = False Then
			DebugTimer ""& PLUGIN_CODE &" table creating"
    		
    		Conn.Execute("SET NAMES utf8mb4;") 
    		Conn.Execute("SET FOREIGN_KEY_CHECKS = 0;") 
    		
    		Conn.Execute("DROP TABLE IF EXISTS `"& PluginTableName &"`")

    		q="CREATE TABLE `"& PluginTableName &"` ( "
    		q=q+"  `ID` int(11) NOT NULL AUTO_INCREMENT, "
    		q=q+"  `MODULE` varchar(255) DEFAULT NULL UNIQUE, "
    		q=q+"  `ORDER` int(9) DEFAULT 0, "
    		q=q+"  `STATUS` int(1) DEFAULT 1, "
    		q=q+"  PRIMARY KEY (`ID`), "
    		q=q+"  KEY `IND1` (`MODULE`) "
    		q=q+") ENGINE=MyISAM DEFAULT CHARSET=utf8; "
			Conn.Execute(q)

    		Conn.Execute("SET FOREIGN_KEY_CHECKS = 1;") 

			' Create Log
			'------------------------------
    		Call PanelLog(""& PLUGIN_CODE &" için database tablosu oluşturuldu", 0, ""& PLUGIN_CODE &"", 0)

			' Register Settings
			'------------------------------
			DebugTimer ""& PLUGIN_CODE &" class_register() End"
    	End If

		' Register Settings
		'------------------------------
		a=GetSettings("PLUGIN:"& PLUGIN_CODE &"", PLUGIN_CODE)
		a=GetSettings(""&PLUGIN_CODE&"_PLUGIN_NAME", PLUGIN_NAME)
		a=GetSettings(""&PLUGIN_CODE&"_CLASS", "Rabbit_Module_Extender_Plugin")
		a=GetSettings(""&PLUGIN_CODE&"_REGISTERED", ""& Now() &"")
		a=GetSettings(""&PLUGIN_CODE&"_CODENO", "63")
		a=GetSettings(""&PLUGIN_CODE&"_ACTIVE", "1")
		a=GetSettings(""&PLUGIN_CODE&"_FOLDER", PLUGIN_FOLDER_NAME)

		' Register Settings
		'------------------------------
		DebugTimer ""& PLUGIN_CODE &" class_register() End"
	End Property
	'---------------------------------------------------------------
	' Register Class
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' Plugin Admin Panel Extention
	'---------------------------------------------------------------
	Public sub LoadPanel()
		'--------------------------------------------------------
		' Sub Page 
		'--------------------------------------------------------
		If Query.Data("Page") = "SHOW:LoadedModules" Then
			Call PluginPage("Header")

			With Response 
				.Write "<div class=""table-responsive"">"
				.Write "	<table class=""table table-striped table-bordered"">"
				.Write "		<thead>"
				.Write "			<tr>"
				.Write "				<th>Sıra</th>"
				.Write "				<th>Modül</th>"
				.Write "				<th>Durum</th>"
				.Write "				<th>İşlem</th>"
				.Write "			</tr>"
				.Write "		</thead>"
				.Write "		<tbody>"
				Set Siteler = Conn.Execute("SELECT * FROM `tbl_plugin_"& PLUGIN_DB_NAME &"` ORDER BY `ORDER` ASC")
				If Siteler.Eof Then 
				    Response.Write "<tr>"
				        Response.Write "<td colspan=""4"" align=""center"">Modül Bulunamadı</td>"
				    Response.Write "</tr>"
				End If
				Do While Not Siteler.Eof
				.Write "			<tr>"
				.Write "				<td>"& Siteler("ORDER") &"</td>"
				.Write "				<td>"& Siteler("MODULE") &"</td>"
				.Write "				<td>"& PLUGIN_STATUS_TEXT( Siteler("STATUS") ) &"</td>"
				.Write "				<td align=""right"">"
				' .Write "					<a data-ajax=""true"" data-remove=""tr"" href=""/panel/ajax.asp?Cmd=PluginSettings&PluginName="& PLUGIN_CODE &"&Page=RemoveLog&RecID="& Siteler("ID") &""" class=""btn btn-sm btn-danger"">"
				' .Write "						Sil"
				' .Write "					</a>"
				.Write "				</td>"
				.Write "			</tr>"
				Siteler.MoveNext : Loop
				Siteler.Close : Set Siteler = Nothing
				.Write "		</tbody>"
				.Write "	</table>"
				.Write "</div>"
			End With

			Call PluginPage("Footer")
			Call SystemTeardown("destroy")
		End If

		'--------------------------------------------------------
		' Main Page
		'--------------------------------------------------------
		With Response
			'------------------------------------------------------------------------------------------
				PLUGIN_PANEL_MASTER_HEADER This()
			'------------------------------------------------------------------------------------------
			' .Write "<div class=""row"">"
			' .Write "    <div class=""col-lg-6 col-sm-12"">"
			' .Write 			QuickSettings("select", ""& PLUGIN_CODE &"_OPTION_1", "Buraya Title", "0#Seçenek 1|1#Seçenek 2|2#Seçenek 3", TO_DB)
			' .Write "    </div>"
			' .Write "    <div class=""col-lg-6 col-sm-12"">"
			' .Write 			QuickSettings("number", ""& PLUGIN_CODE &"_OPTION_2", "Buraya Title", "", TO_DB)
			' .Write "    </div>"
			' .Write "    <div class=""col-lg-12 col-sm-12"">"
			' .Write 			QuickSettings("tag", ""& PLUGIN_CODE &"_OPTION_3", "Buraya Title", "", TO_DB)
			' .Write "    </div>"
			' .Write "</div>"

			.Write "<div class=""row"">"
			.Write "    <div class=""col-lg-12 col-sm-12"">"
			.Write "        <a open-iframe href=""ajax.asp?Cmd=PluginSettings&PluginName="& PLUGIN_CODE &"&Page=SHOW:LoadedModules"" class=""btn btn-sm btn-primary"">"
			.Write "        	Yüklü Modülleri Göster"
			.Write "        </a>"
			' .Write "        <a open-iframe href=""ajax.asp?Cmd=PluginSettings&PluginName="& PLUGIN_CODE &"&Page=DELETE:CachedFiles"" class=""btn btn-sm btn-danger"">"
			' .Write "        	Tüm Önbelleği Temizle"
			' .Write "        </a>"
			.Write "    </div>"
			.Write "</div>"
		End With
	End Sub
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' Class First Init
	'---------------------------------------------------------------
	Private Sub class_initialize()
    	'-------------------------------------------------------------------------------------
    	' PluginTemplate Main Variables
    	'-------------------------------------------------------------------------------------
    	PLUGIN_NAME 			= "Rabbit Module Extender Plugin"
    	PLUGIN_CODE  			= "RABBIT_MODULE_EXTENDER"
    	PLUGIN_DB_NAME 			= "module_extender"
    	PLUGIN_VERSION 			= "1.0.0"
    	PLUGIN_CREDITS 			= "@badursun Anthony Burak DURSUN"
    	PLUGIN_GIT 				= "https://github.com/RabbitCMS-Hub/Rabbit-Module-Extender-Plugin"
    	PLUGIN_DEV_URL 			= "https://adjans.com.tr"
    	PLUGIN_ICON 			= "zmdi-pin-help"
    	PLUGIN_REMOVABLE 		= True
    	PLUGIN_FILES_ROOT 		= PLUGIN_VIRTUAL_FOLDER(This)
    	'-------------------------------------------------------------------------------------
    	' PluginTemplate Main Variables
    	'-------------------------------------------------------------------------------------

		MODULE_ROOT 			= "/content/vendor/"
		LOADED_MODULES 			= ""
		MODULE_STATUS 			= Cint( GetSettings(""& PLUGIN_CODE &"_ACTIVE", "1") )

		ORDER_JS 				= 60
		ORDER_CSS 				= 80

    	'-------------------------------------------------------------------------------------
    	' PluginTemplate Register App
    	'-------------------------------------------------------------------------------------
    	class_register()
	End Sub
	'---------------------------------------------------------------
	' Class First Init
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' Class Terminate
	'---------------------------------------------------------------
	Private sub class_terminate()

	End Sub
	'---------------------------------------------------------------
	' Class Terminate
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' Plugin Defines
	'---------------------------------------------------------------
	Public Property Get PluginCode() 		: PluginCode = PLUGIN_CODE 					: End Property
	Public Property Get PluginName() 		: PluginName = PLUGIN_NAME 					: End Property
	Public Property Get PluginVersion() 	: PluginVersion = PLUGIN_VERSION 			: End Property
	Public Property Get PluginGit() 		: PluginGit = PLUGIN_GIT 					: End Property
	Public Property Get PluginDevURL() 		: PluginDevURL = PLUGIN_DEV_URL 			: End Property
	Public Property Get PluginFolder() 		: PluginFolder = PLUGIN_FILES_ROOT 			: End Property
	Public Property Get PluginIcon() 		: PluginIcon = PLUGIN_ICON 					: End Property
	Public Property Get PluginRemovable() 	: PluginRemovable = PLUGIN_REMOVABLE 		: End Property
	Public Property Get PluginCredits() 	: PluginCredits = PLUGIN_CREDITS 			: End Property
	Public Property Get PluginRoot() 		: PluginRoot = PLUGIN_ROOT 					: End Property
	Public Property Get PluginFolderName() 	: PluginFolderName = PLUGIN_FOLDER_NAME 	: End Property
	Public Property Get PluginDBTable() 	: PluginDBTable = IIf(Len(PLUGIN_DB_NAME)>2, "tbl_plugin_"&PLUGIN_DB_NAME, "") 	: End Property

	Private Property Get This()
		This = Array(PluginCode, PluginName, PluginVersion, PluginGit, PluginDevURL, PluginFolder, PluginIcon, PluginRemovable, PluginCredits, PluginRoot, PluginFolderName, PluginDBTable )
	End Property
	'---------------------------------------------------------------
	' Plugin Defines
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Property Let Load(vModuleName)
		If MODULE_STATUS = 0 Then 
			Exit Property
		End If

		If Len(vModuleName) < 1 Then 
			Exit Property
		End If

		Dim ModuleVendorRoot, ModuleVendorPath
			ModuleVendorPath = MODULE_ROOT&vModuleName&"/"
			ModuleVendorRoot = Server.Mappath(ModuleVendorPath)

		If IsFolderExist(ModuleVendorRoot) Then
			AddToDB vModuleName, 0

	        Set ModuleFiles = Server.CreateObject("Scripting.FileSystemObject")
                Set FileList = ModuleFiles.GetFolder( ModuleVendorRoot )
                    For Each x In FileList.files
                        If LCase(Right(x.Name, 3)) = "css" Then 
                        	If ORDER_CSS < 60 AND ORDER_CSS > 79 Then 
                        		Exit For
                        	End If
                        	Cms.CSS ORDER_CSS, Query.Asset(ModuleVendorPath + x.Name)
                        	ORDER_CSS=ORDER_CSS+1
                        	' Cms.AddCSS ModuleVendorPath & x.Name , vModuleName
                        End If
                        If LCase(Right(x.Name, 2)) = "js" Then
                        	If ORDER_JS < 80 AND ORDER_JS > 99 Then 
                        		Exit For
                        	End If 
                        	Cms.JS ORDER_JS, Query.Asset(ModuleVendorPath + x.Name)
                        	ORDER_JS=ORDER_JS+1
                        	'Cms.AddJS ModuleVendorPath & x.Name , vModuleName
                        End If
                    Next
                Set FileList = Nothing
	        Set ModuleFiles = Nothing
		Else
			Exit Property
		End If

	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Private Sub AddToDB(Module, OrderNo)
		Conn.Execute("INSERT IGNORE INTO `tbl_plugin_"& PLUGIN_DB_NAME &"`(`MODULE`, `ORDER`) VALUES('"& Module &"','"& OrderNo &"')")
	End Sub
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
End Class 
%>
