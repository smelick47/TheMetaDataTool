<%-- The following 4 lines are ASP.NET directives needed when using SharePoint components --%>

<%@ Page Inherits="Microsoft.SharePoint.WebPartPages.WebPartPage, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" MasterPageFile="~masterurl/default.master" Language="C#" %>

<%@ Register TagPrefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>

<%-- The markup and script in the following Content element will be placed in the <head> of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderAdditionalPageHead" runat="server">
    <script type="text/javascript" src="../Scripts/ScriptBase.js"></script>
</asp:Content>

<%-- The markup in the following Content element will be placed in the TitleArea of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea" runat="server">
    Home
</asp:Content>

<%-- The markup and script in the following Content element will be placed in the <body> of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderMain" runat="server">

    <div id="tabs" class="easyui-tabs" style="height: auto; min-width: 600px">

        <div title="<span class='tt-inner'><img src='../Images/home32.png'/><br>Home</span>" data-options="closable:false">

            <h2 class="ms-webpart-titleText">Available Manage Metadata Services</h2>
            <div id="mms-items" title="Click here to select" style="padding: 5px; border: 1px solid #ddd;">
            </div>
        </div>
        <div title="<span class='tt-inner'><img src='../Images/properties32.png'/><br>Properties</span>">
            <div id="tbproperties">
                <table id="pgMetaData" class="easyui-propertygrid" style="width: 600px"></table>
            </div>
        </div>
        <div title="<span class='tt-inner'><img src='../Images/data32.png'/><br>Data</span>" data-options="closable:false" style="min-height: 450px">
            <div id="tbData">
                <div id="tb-data-toolbar-data" style="margin: 5px 0px 5px 0px; padding: 5px; border: 1px solid #ddd; min-height: 28px;">
                    <a id="btn-loadtem">Refresh</a>
                    <a id="btn-saveprop">Save Properties</a>
                </div>
                <div id="TreeCategories" class="tabMetaDataTreeSection">
                    <ul id="metaDataTree" class="easyui-tree"></ul>

                    <div id="metaDataTree-context-store" style="width: 120px;">
                        <div onclick="app.Default.contextMenuTermStoreAddGroup()" data-options="iconCls:'icon-add'">New Group</div>
                    </div>

                    <div id="metaDataTree-context-group" style="width: 120px;">
                        <div onclick="app.Default.contextMenuGroupAddTermSet()" data-options="iconCls:'icon-add'">New Term Set</div>
                        <div onclick="app.Default.contextMenuGroupDeleteGroup()" data-options="iconCls:'icon-delete'">Delete Group</div>
                    </div>

                    <div id="metaDataTree-context-termset" style="width: 120px;">
                        <div onclick="app.Default.contextMenuTermSetCreateTerm()" data-options="iconCls:'icon-add'">New Term</div>
                        <div onclick="app.Default.contextMenuTermSetDeleteTermSet()" data-options="iconCls:'icon-delete'">Delete Term Set</div>
                    </div>

                    <div id="metaDataTree-context-term" style="width: 120px;">
                        <div onclick="app.Default.contextMenuTermCreateTerm()" data-options="iconCls:'icon-add'">New Term</div>
                        <div onclick="app.Default.contextMenuTermDeleteTerm()" data-options="iconCls:'icon-delete'">Delete Term</div>
                    </div>

                </div>

                <div id="termProperties" class="termMetaDataProperties">
                    <table id="termMetaData" class="easyui-propertygrid" style="width: 450px"></table>
                </div>
                <div id="termLanguage" class="termLanguageSection">
                    <div style="margin: 5px 0px 5px 0px; padding: 5px; border: 1px solid #ddd">
                        Language: &nbsp;
                        <select id="slcttermlanguages" onchange="app.Default.mmstermLanguageData()">
                        </select>
                        <a id="btn-lbl-add" title="Add Label"></a>
                        <a id="btn-lbl-add-cp" title="Add Custom Property"></a>
                        <a id="btn-lbl-del" title="Delete Selected"></a>

                        <div id="dlg-add-Labels" class="dlg">
                            <label>Label :</label>
                            <input type="text" name="dlg-add-Labels-label" />
                        </div>

                        <div id="dlg-add-Properties" class="dlg">
                            <table>
                                <tr>
                                    <td>Type :</td>
                                    <td>
                                        <label>Shared Property</label>
                                        <input type="radio" name="dlg-add-Properties-type" checked="checked" value="true" />
                                        <label>Local Property </label>
                                        <input type="radio" name="dlg-add-Properties-type" value="false" />
                                    </td>
                                </tr>
                                <tr>
                                    <td>Name :</td>
                                    <td>
                                        <input type="text" name="dlg-add-Properties-name" /></td>
                                </tr>
                                <tr>
                                    <td>Value :</td>
                                    <td>
                                        <input type="text" name="dlg-add-Properties-value" /></td>
                                </tr>
                            </table>
                        </div>

                    </div>
                    <table id="termLanguageData" class="easyui-propertygrid" style="width: 350px"></table>
                </div>
            </div>
        </div>
        <div title="<span class='tt-inner'><img src='../Images/importexport32.png'/><br>Import & Export </span>" data-options="closable:false">
            <div id="tbimport">

                <div id="tb-exportimport-toolbar-export" style="margin: 5px 0px 5px 0px; padding: 5px; border: 1px solid #ddd; min-height: 28px;">
                    <a id="btn-refresh-tab-importexport">Refresh</a>
                    <a id="btn-export">Export</a>
                    <a id="btn-import">Import</a>
                </div>

                <div class="tabImportExporttExportSection">
                    <h2 class="ms-webpart-titleText">Exported Data</h2>
                    <table id="exportGrid" style="width: 400px; height: 320px">
                        <thead>
                            <tr>
                                <th field="tmp" width="250">Name</th>
                                <th field="dt" width="90">Date</th>
                            </tr>
                        </thead>
                    </table>
                </div>

                <div class="exportPreviewSection">
                    <ul id="exportTreePreview" class="easyui-tree"></ul>
                </div>

                <div id="dlg-import-terms" class="dlg">
                    <div class="status"></div>
                    <span>Import Options</span>
                    <br />
                    <%--    <label>Overide if exist</label><input type="radio" name="dlg-import-terms-options" value="1" />--%>
                    <label>Fail if exist </label>
                    <input title="Import will be canceld if it found any existing group." type="radio" name="dlg-import-terms-options" checked="checked" value="2" />
                    <label>Skip if exist </label>
                    <input title="Existing groups will be skiped." type="radio" name="dlg-import-terms-options" value="3" />
                    <br />
                    <h4>Select groups form import</h4>
                    <ul id="ImporttermaDataTree" class="easyui-tree"></ul>
                </div>

                <div id="ExportData" class="tabImportExporttExportSection" style="display: none">
                    <div>
                        Exported Data From File : &nbsp;
                        <br />
                        <select id="slcexportData" onchange="app.Default.loadExportData()"></select>

                        <ul id="ExportedmetaDataTree" class="easyui-tree"></ul>
                    </div>



                </div>

            </div>
        </div>

        <div title="<span class='tt-inner'><img src='../Images/templates32.png'/><br>Templates </span>" data-options="closable:false">
            <div id="tbtemplates">
                <div id="tb-templates-toolbar" style="margin: 5px 0px 5px 0px; padding: 5px; border: 1px solid #ddd; min-height: 28px;">
                    <a id="btn-template-refresh">Refresh</a>
                    <a id="btn-saveasatemplate">Save as a Template</a>
                </div>
                <div id="TemplateTree" class="tabTemplateTreeSection">
                    <div id="templateTree-context" style="width: 180px">
                        <div onclick="app.Default.TabTemplate.ExportAsTemplate();" data-options="iconCls:'icon-export'">Save as Template</div>
                        <div onclick="app.Default.TabTemplate.loadDlgApply();" data-options="iconCls:'icon-import'">Apply Template</div>
                    </div>
                    <div id="templateTree-context-root" style="width: 180px">
                        <div onclick="app.Default.TabTemplate.loadDlgApply();" data-options="iconCls:'icon-import'">Apply Template</div>
                    </div>
                    <ul id="templateTree" class="easyui-tree"></ul>
                </div>
                <div class="templateSection">
                    <h2 class="ms-webpart-titleText">Available Templates</h2>
                    <table id="templateGrid" style="width: 350px; height: 320px">
                        <thead>
                            <tr>
                                <th field="tmp" width="250">Name</th>
                                <th field="typ" width="90">Template Type</th>
                            </tr>
                        </thead>
                    </table>
                </div>

                <div class="templatePreviewSection">
                    <ul id="templateTreePreview" class="easyui-tree"></ul>
                </div>

            </div>

            <div id="dlg-apply-Templates" class="dlg">
                <table id="from-apply-Templates">
                    <tr>
                        <td>Templates :</td>
                        <td>
                            <select id="templatesToAplly" name="templatesToAplly"></select>
                        </td>
                    </tr>
                    <%--                    <tr>
                        <td>Enter Name :</td>
                        <td>
                            <input id="dlgTemplateName" name="templateName" />
                        </td>
                    </tr>--%>
                </table>
            </div>

        </div>
    </div>

    <script type="text/javascript">
        $$.SharePointReady(function () {
            $(document).ready(function () {

                app.Default.onLoad();

                $('#metaDataTree').tree({
                    onBeforeExpand: function (node) {
                        return app.Default.loadChildData(node, "#metaDataTree");
                    },
                    onClick: function (node) {
                        app.Default.mmsTreeNodeClick(node);
                    }
                });

                $('#templateTree').tree({
                    onBeforeExpand: function (node) {
                        return app.Default.loadChildData(node, "#templateTree");
                    }
                });

                $("a#btn-loadtem").click(function () {
                    app.Default.loadMetadata();
                    $('#termMetaData').propertygrid({ data: new mapers.PropertyData() });
                    $('#termLanguageData').propertygrid({ data: new mapers.PropertyData() });
                });

                $("a#btn-saveprop").on("click", app.Default.saveChangestermProperties);

                // tab import export
                $("a#btn-export").on("click", app.Default.TabImportExport.exportTermStore);
                $("a#btn-import").on("click", app.Default.TabImportExport.openImportDialog);
                $("a#btn-refresh-tab-importexport").on("click", app.Default.TabImportExport.loadTab);

                $("a#btn-lbl-add").on("click", app.Default.addLabels);
                $("a#btn-lbl-del").on("click", app.Default.deleteLanguageData);
                $("a#btn-lbl-add-cp").on("click", app.Default.addCustomProperty);

                $("a#btn-template-refresh").on("click", app.Default.TabTemplate.loadTab);
                $("a#btn-saveasatemplate").on("click", app.Default.TabTemplate.ExportAsTemplate);


            });
        }, function () { handeler.waitUILongProcess(); });
    </script>

</asp:Content>
