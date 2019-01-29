/// <reference path="../scriptbase.js" />

(function (ei, $, undefined) {

    ei.Options = {
        Export: {
            TermProperties: false, // 
            TermLables: false //
        }

    };

    ei.getObjectData = function () {
        return {
            "tm": new mapers.Template(),
            "cc": 0, "rc": 0, "fc": 0, "lg": "", "Options": ei.Options,
            log: function (log) { this.lg += ";" + log }
        };
    };

    ei.mapGroupObject = function (cu) {
        var t = new mapers.G();
        t.d = cu.get_description();
        t.i = cu.get_id().toString();
        t.x = cu.get_name();

        return t;
    };

    ei.mapTermSetObject = function (cu) {
        var t = new mapers.S();
        t.d = cu.get_description();
        t.i = cu.get_id().toString();
        t.x = cu.get_name();

        return t;
    };

    ei.mapTermObject = function (cu) {
        var t = new mapers.T();
        t.d = cu.get_description();
        t.i = cu.get_id().toString();
        t.x = cu.get_name();

        return t;
    };

    ei.mapLabelObject = function (cu) {
        var lbl = new mapers.OL();
        lbl.c = cu.get_language();
        lbl.v = cu.get_value();
        lbl.d = cu.get_isDefaultForLanguage();

        return lbl;
    };

    //i = ref Group, ts = ref TermSet, This is a reference to a object
    ei.exportGetTermsForTermSet = function (refG, refTs, data) {

        //data = ei.getObjectData();
        // ts = new mapers.S();

        var mms = new spa.Metadata();

        var i = refG.i; // Current Group Id
        var j = refTs.i; // Current Term Set ID

        data.cc++;
        mms.getTermsAsync(data.tm.Data.i, i, j)
            .done(function (terms) {

                data.rc++;
                var enu = terms.getEnumerator();
                while (enu.moveNext()) {
                    var cu = enu.get_current();
                    var t = ei.mapTermObject(cu);
                    refTs.t.push(t);

                    // get other values
                    ei.exportTermData(refTs.t[refTs.t.length - 1], refTs, refG, data);
                    if (cu.get_termsCount() > 0) {
                        ei.exportGetChildTerms(refTs.t[refTs.t.length - 1], refG, refTs, data);
                    }

                }
            })
            .fail(function (b) {
                data.fc++;
                data.log(" Term faild." + b);
            });
    };

    // iterate objdata.tm.Data.g and get term sets and terms
    ei.exportGetTermSetsandTerms = function (data) {

        var mms = new spa.Metadata();

        $.each(data.tm.Data.g, function (i, ig) {
            data.cc++;

            mms.getallTermSetsbyGroupAsync(data.tm.Data.i, ig.i)
                .done(function (TermSets) {
                    data.rc++
                    var enu = TermSets.getEnumerator();
                    while (enu.moveNext()) {
                        var cu = enu.get_current();
                        var s = ei.mapTermSetObject(cu);
                        data.tm.Data.g[i].s.push(s);
                        ei.exportGetTermsForTermSet(data.tm.Data.g[i], data.tm.Data.g[i].s[data.tm.Data.g[i].s.length - 1], data);
                    }
                }).fail(function (b) {
                    data.fc++;
                    data.log(" Term Sets faild." + b);
                });
        });

    }

    // reference parent term , refrence group reference termset to javascript object
    ei.exportGetChildTerms = function (refParentterm, refGroup, refTS, data) {

        //data = objdata; //TODO delete this is for intelicence
        // refGroup = new mapers.G();
        // refTS = new mapers.S();
        //refParentterm = new mapers.T();

        var i = refGroup.i; // Current Group Id
        var j = refTS.i; // Current Term Set ID

        var mms = new spa.Metadata();

        data.cc++;
        mms.getChildTermsAsync(data.tm.Data.i, i, j, refParentterm.i)
            .done(function (cterms) {
                data.rc++;
                var enu = cterms.getEnumerator();
                while (enu.moveNext()) {
                    var cu = enu.get_current();
                    var t = ei.mapTermObject(cu);
                    refParentterm.t.push(t);

                    // get other values
                    ei.exportTermData(refParentterm.t[refParentterm.t.length - 1], refTS, refGroup, data);
                    if (cu.get_termsCount() > 0) {
                        ei.exportGetChildTerms(refParentterm.t[refParentterm.t.length - 1], refGroup, refTS, data);
                    }
                }
            })
            .fail(function (b) {
                data.fc++
                data.log(" Child Term failed " + b);
            });
    };

    // Group the term export for options
    ei.exportTermData = function (refTerm, refTS, refGroup, data) {

        //data = ei.getObjectData();

        // Export shared and local properties
        if (data.Options.Export.TermProperties) {
            ei.exportTermLocalSharedProperties(refTerm, refTS, refGroup, data);
        }

        // export other language labes
        if (data.Options.Export.TermLables) {
            ei.exportTermLables(refTerm, refTS, refGroup, data);
        }
    };

    ei.exportTermLocalSharedProperties = function (refTerm, refTS, refGroup, data) {

        //data = objdata; //TODO delete this is for intelicence
        //refGroup = new mapers.G();
        //refTS = new mapers.S();
        //refTerm = new mapers.T();

        var mms = new spa.Metadata();

        data.cc++
        mms.getTermLocalSharedPropertiesByTermIdAsync(data.tm.Data.i, refGroup.i, refTS.i, refTerm.i).done(function (lp, sp) {
            data.rc++;
            $.each(lp, function (k, v) {
                var p = new mapers.P()
                p.n = k;
                p.v = v;
                refTerm.w.push(p);
            });

            $.each(sp, function (k, v) {
                var p = new mapers.SP()
                p.n = k;
                p.v = v;
                refTerm.y.push(p);
            });
        }).fail(function (b) {
            data.fc++
            data.log(" Term properties failed. " + b);
        });
    };

    ei.exportTermLables = function (refTerm, refTS, refGroup, data) {

        //var data = ei.getObjectData(); //TODO delete this
        //var refGroup = new mapers.G();
        //var refTS = new mapers.S();
        //var refTerm = new mapers.T();

        var mms = new spa.Metadata();

        $.each(data.tm.l, function (i, lan) {

            data.cc++;
            mms.getTermByIdAsyncWithLabels(data.tm.Data.i, refGroup.i, refTS.i, refTerm.i, lan)
                .done(function (lables) {
                    data.rc++;
                    var enu = lables.getEnumerator();
                    while (enu.moveNext()) {
                        var cu = enu.get_current();

                        if (lan == data.tm.Data.dl && cu.get_isDefaultForLanguage()) {
                            // This is captured as name
                        }
                        else {
                            var lbl = ei.mapLabelObject(cu);
                            refTerm.z.push(lbl);
                        }
                    }
                }).fail(function (b) {
                    objdata.fc++
                    data.log(" Term lable failed. " + refTerm.x + ":" + b);
                });
        });

    };

})(window.ei = window.ei || {}, jQuery);

(function (app, $, undefined) {

    // loadSuiteBar();
    cssLoads();

    // check first time and assign list permisions to all
    function appInstalation(done, failed) {
        var op = new spa.Operations();
        op.getListItemByIdAsync(defs.L.Configs.Nm, 1).done(function (itm) {
            // New Instalation
            if (itm.get_item('ConfigItem') == 0) {
                var userobj = spa.web.ensureUser("c:0(.s|true");
                var role = SP.RoleDefinitionBindingCollection.newObject(spa.ctx);
                role.add(spa.web.get_roleDefinitions().getByType(SP.RoleType.contributor));

                var oList = spa.web.get_lists().getByTitle(defs.L.Export.Nm);
                oList.breakRoleInheritance(false, true);
                oList.get_roleAssignments().add(userobj, role);

                var oCat = spa.web.get_lists().getByTitle(defs.L.LogData.Nm);
                oCat.breakRoleInheritance(false, true);
                oCat.get_roleAssignments().add(userobj, role);

                spa.executeCtxAsync()
                    .done(function () {
                        op.updateListColumnByIdAsync(defs.L.Configs.Nm, 1, defs.L.Configs.C.ConfigItem, 1)
                            .done(function () { done(); })
                            .fail(function () { failed(); });
                    })
                    .fail(function () { failed(); });
            } else { done(); }
        }).fail(function () { failed(); });
    }

    function cssLoads() {
        // $("h1#pageTitle").remove();
    }

    function loadSuiteBar() {
        var p = new spa.Permisions();
        p.doesUserHaveManagewebAsync().done(function () {
            handeler.AddtoSuiteBar("app-Settings-ico", "Settings", handeler.getImageUrl("settings.png"), "", "handeler.goTo('pages/_app/Appsettings.aspx');");
        });
    };

    function getmmsid() {
        return $("span#DeltaPlaceHolderPageTitleInTitleArea").attr("data-mmsid");
    };

    function getmmsname() {
        return $("span#DeltaPlaceHolderPageTitleInTitleArea").text();
    };

    function getDefailtLanguage() {
        return $("span#DeltaPlaceHolderPageTitleInTitleArea").attr("data-dl");
    };

    function getMMSLanguages() {

        var ret = [];

        $.each(JSON.parse($("span#DeltaPlaceHolderPageTitleInTitleArea").attr("data-l")), function (i, ele) {
            ret.push(ele);
        });

        return ret;
    };

    // Map TermStore Object to Tree Node
    function itemToNodeExportTree(i) {
        var itm = new mapers.TreeItem();

        itm.Id = i.i;
        itm.text = i.x;

        if (i[mapers.TreeItem.Attributes.isGroup]) {
            itm.iconCls = 'icon-mmgroup';
        }
        else if (i[mapers.TreeItem.Attributes.isTermSet]) {
            itm.iconCls = 'icon-mmTermSet';
        }
        else if (i[mapers.TreeItem.Attributes.isTerm]) {
            itm.iconCls = 'icon-mmTerm';
        };

        itm.state = mapers.TreeItem.State.Close;
        return itm;
    }

    // Map TermStore Object to Tree Node Child items
    function loadChildTerms(treeNode, parentTerm) {
        parentTerm.forEach(function (ct) {
            // define as Term 

            ct[mapers.TreeItem.Attributes.isTerm] = true;
            var itmct = itemToNodeExportTree(ct);
            if (ct.t.length > 0) {
                loadChildTerms(itmct, ct.t);
            }
            else {
                itmct.state = mapers.TreeItem.State.Open;
            }
            treeNode.children.push(itmct);
        })
    }

    // Map TermStore Object to Tree Node and generate the tree
    function loadExportMetaDataToTree(tergetTreeSelector, objTermStore) {

        var treeData = [];

        objTermStore.g.forEach(function (G) {

            // define as Group
            G[mapers.TreeItem.Attributes.isGroup] = true;
            var itmg = itemToNodeExportTree(G);

            if (G.s.length > 0) {
                G.s.forEach(function (TS) {
                    // define as Term Set

                    TS[mapers.TreeItem.Attributes.isTermSet] = true;
                    var itms = itemToNodeExportTree(TS);

                    if (TS.t.length > 0) {
                        TS.t.forEach(function (T) {
                            // define as Term 

                            T[mapers.TreeItem.Attributes.isTerm] = true;
                            var itmt = itemToNodeExportTree(T);

                            if (T.t.length > 0) {
                                loadChildTerms(itmt, T.t);
                            }
                            else {
                                itmt.state = mapers.TreeItem.State.Open;
                            }

                            itms.children.push(itmt);
                        });
                    }
                    else {
                        itms.state = mapers.TreeItem.State.Open;
                    }

                    itmg.children.push(itms);
                });
            }
            else {
                itmg.state = mapers.TreeItem.State.Open;
            }

            treeData.push(itmg);
        });

        var rootData = [];
        var root = new mapers.TreeItem();
        root.Id = objTermStore.i;
        root.text = objTermStore.n;
        root.state = mapers.TreeItem.State.Open;
        root.children = treeData;
        root.attributes = { "Store": true };
        rootData.push(root);

        $(tergetTreeSelector).tree({
            data: rootData,
            checkbox: false
        });

    };

    //Pager for tempalte preview grid 
    function pagerFilterTabTemplate(d) {
        if (typeof d.length == 'number') {    // is array
            data = {
                total: d.length,
                rows: d
            }
        }
        var dg = $(this);
        var opts = dg.datagrid('options');
        //opts.pageSize = 8; // Bug Fix , pase site not set

        var pager = dg.datagrid('getPager');
        //var t = data.total % opts.pageSize;
        //pager.pagination('refresh', { total: t });

        pager.pagination({
            onSelectPage: function (pageNum, pageSize) {
                opts.pageNumber = pageNum;
                opts.pageSize = pageSize;
                pager.pagination('refresh', {
                    pageNumber: pageNum,
                    pageSize: pageSize
                });
                dg.datagrid('loadData', data);
            }
        });
        if (!data.originalRows) {
            data.originalRows = data.rows;
        }
        var start = (opts.pageNumber - 1) * parseInt(opts.pageSize);
        var end = start + parseInt(opts.pageSize);
        data.rows = (data.originalRows.slice(start, end));
        return data;
    };

    app.Default = (function () {

        var obj = {};

        function disabletabs(disable) {
            if (disable) {
                $('#tabs').tabs('disableTab', 1);
                $('#tabs').tabs('disableTab', 2);
                $('#tabs').tabs('disableTab', 3);
                $('#tabs').tabs('disableTab', 4);

            }
            else {
                $('#tabs').tabs('enableTab', 1);
                $('#tabs').tabs('enableTab', 2);
                $('#tabs').tabs('enableTab', 3);
                $('#tabs').tabs('enableTab', 4);
            }
        }

        function controlInt() {

            $('#tabs').tabs({
                border: false, tabHeight: 55, tabPosition: 'top',
                onSelect: function (title, index) {
                    switch (index) {
                        case 1:
                            break;
                        case 2:
                            break;
                        case 3:
                            app.Default.TabImportExport.loadTab();
                            break;
                        case 4:
                            app.Default.TabTemplate.loadTab();
                            break;
                        default:
                    }
                }
            });

            disabletabs(true);

            $('a#btn-loadtem').linkbutton({ plain: true, iconCls: 'icon-reload' });
            $('a#btn-saveprop').linkbutton({ plain: true, iconCls: 'icon-save' });

            // tab import export
            $('a#btn-export').linkbutton({ plain: true, iconCls: 'icon-export' });
            $('a#btn-import').linkbutton({ plain: true, iconCls: 'icon-import' });
            $('a#btn-refresh-tab-importexport').linkbutton({ plain: true, iconCls: 'icon-reload' });

            $('a#btn-lbl-add').linkbutton({ plain: true, iconCls: 'icon-add' });
            $('a#btn-lbl-add-cp').linkbutton({ plain: true, iconCls: 'icon-add' });
            $('a#btn-lbl-del').linkbutton({ plain: true, iconCls: 'icon-remove' });

            //tab Template
            $('a#btn-template-refresh').linkbutton({ plain: true, iconCls: 'icon-reload' });
            $('a#btn-saveasatemplate').linkbutton({ plain: true, iconCls: 'icon-save' });
            $('a#btn-tmpl-all').linkbutton({ plain: true, toggle: true, group: 'tmpltypes' });
            $('a#btn-tmpl-store').linkbutton({ plain: true, toggle: true, group: 'tmpltypes' });
            $('a#btn-tmpl-group').linkbutton({ plain: true, toggle: true, group: 'tmpltypes' });
            $('a#btn-tmpl-set').linkbutton({ plain: true, toggle: true, group: 'tmpltypes' });
            $('a#btn-tmpl-term').linkbutton({ plain: true, toggle: true, group: 'tmpltypes' });


            $('#templateGrid').datagrid({
                fitColumns: true, autoRowHeight: false,
                pagination: true, pageSize: 7, pageList: [7, 14, 21, 28],
                singleSelect: true,
                onClickRow: app.Default.TabTemplate.clickTemplateGrid
            });

            $('#exportGrid').datagrid({
                fitColumns: true, autoRowHeight: false,
                pagination: true, pageSize: 7, pageList: [7, 14, 21, 28],
                singleSelect: true,
                onClickRow: app.Default.TabImportExport.clickExportGrid
            });

            // hide refresh button
            $('#templateGrid').datagrid('getPager').pagination({ 'showRefresh': false });
            $('#templateTree-context').menu();
            $('#templateTree-context-root').menu();


            $('#metaDataTree-context-store').menu();
            $('#metaDataTree-context-group').menu();
            $('#metaDataTree-context-termset').menu();
            $('#metaDataTree-context-term').menu();


            $('div#dlg-add-Labels').dialog({
                title: 'Add Synonyms & Abbreviations',
                width: '250',
                height: 'auto',
                modal: true,
                closable: true,
                closed: true,
                toolbar: [{
                    text: 'Save',
                    iconCls: 'icon-save',
                    handler: saveOtherLable
                }]
            });

            $('div#dlg-add-Properties').dialog({
                title: 'Add/Update Custom Property',
                width: '300',
                height: 'auto',
                modal: true,
                closable: true,
                closed: true,
                toolbar: [{
                    text: 'Save',
                    iconCls: 'icon-save',
                    handler: saveCustomProperties
                }]
            });

            $('div#dlg-import-terms').dialog({
                title: 'Import Data',
                width: '400',
                height: '450',
                modal: true,
                closable: true,
                closed: true,
                buttons: [{
                    text: 'Import',
                    handler: app.Default.TabImportExport.importExportedData
                }]
            });

            $('div#dlg-apply-Templates').dialog({
                title: 'Apply Template',
                width: '400',
                height: '200',
                modal: true,
                closable: true,
                closed: true,
                toolbar: [{
                    text: 'Apply Template',
                    iconCls: 'icon-import',
                    handler: app.Default.TabTemplate.applyTemplate
                }]
            });

            $.messager.progress('close');

        }

        function loadDataToTemplateGrid(d) {
            $('#templateGrid').datagrid({ loadFilter: pagerFilterTabTemplate }).datagrid('loadData', d);
        };

        function loadDataToExportGrid(d) {
            $('#exportGrid').datagrid({ loadFilter: pagerFilterTabTemplate }).datagrid('loadData', d);
        };

        function generatemmsitem(cu) {

            $(document).ready(function () {

                var str = '<div class="mms-item"><a href="#" onClick="app.Default.selectmmsItem(this)" id="mms-store" data-name="@data-name" data-dl="@data-dl" data-l="@data-l"  data-mmsid="@data-mmsid"><table><tr><th>Name</th><td>@Name</td></tr><tr><th>Online</th><td>@online</td></tr></table></a></div>';
                str = str.replace("@data-mmsid", cu.get_id());
                str = str.replace("@data-name", cu.get_name());
                str = str.replace("@data-dl", cu.get_defaultLanguage());
                str = str.replace("@data-l", JSON.stringify(cu.get_languages()));
                str = str.replace("@Name", cu.get_name());
                str = str.replace("@online", cu.get_isOnline());
                $("div#mms-items").append(str);
            });

        }

        function groupToTreeNode(group) {
            var itm = new mapers.TreeItem();
            itm.Id = group.get_id();
            itm.text = group.get_name();
            itm.state = mapers.TreeItem.State.Close;
            itm.attributes = { "Group": true };

            if (group.get_isSystemGroup()) {
                itm.iconCls = 'icon-System';
            }
            else if (group.get_isSiteCollectionGroup()) {
                itm.iconCls = 'icon-siteCollection';
            }
            else {
                itm.iconCls = 'icon-mmgroup';
            }

            return itm;
        }

        function termSetToTreeNode(termSet, groupId) {
            var itm = new mapers.TreeItem();
            itm.Id = termSet.get_id();
            itm.text = termSet.get_name();
            itm.state = mapers.TreeItem.State.Close;
            itm.attributes = { "TermSet": true, "GroupID": groupId };
            itm.iconCls = 'icon-mmTermSet';
            return itm;
        }

        function termToTreeNode(term, groupId, termSetId) {
            var itm = new mapers.TreeItem();
            itm.Id = term.get_id();
            itm.text = term.get_name();
            itm.attributes = { "Term": true, "GroupID": groupId, "TermSetID": termSetId };
            if (term.get_termsCount() > 0) {
                itm.state = mapers.TreeItem.State.Close;
            }
            itm.iconCls = 'icon-mmTerm';

            return itm;
        }

        obj.contextMenuTermStoreAddGroup = function () {
            $.messager.prompt('Insert', 'Group Name :', function (r) {
                if (r) {
                    var mms = new spa.Metadata();
                    handeler.loader("div#tb-data-toolbar-data", "loader-tb-data", "tb-data-toolbar-loader");
                    mms.createTermGroupAsync(getmmsid(), r, SP.Guid.newGuid().toString())
                        .fail(function (e) {
                            sputils.flashNotificationError(handeler.parseSPException(e));
                        })
                        .always(function () {
                            app.Default.loadMetadata();
                            handeler.removeloader("#tb-data-toolbar-loader");
                        });
                }
            });
        }

        obj.contextMenuGroupAddTermSet = function () {
            $.messager.prompt('Insert', 'Term Set Name :', function (r) {
                if (r) {
                    var mms = new spa.Metadata();
                    handeler.loader("div#tb-data-toolbar-data", "loader-tb-data", "tb-data-toolbar-loader");

                    var node = $('#metaDataTree').tree('getSelected');

                    mms.createTermSetAsync(getmmsid(), node.Id.toString(), SP.Guid.newGuid().toString(), r, getDefailtLanguage())
                        .fail(function (e) {
                            sputils.flashNotificationError(handeler.parseSPException(e));
                        }).always(function () {
                            app.Default.loadMetadata();
                            handeler.removeloader("#tb-data-toolbar-loader");
                        });
                }
            });
        }

        obj.contextMenuGroupDeleteGroup = function () {
            $.messager.confirm('Confirm', defs.DeleteConfirmation, function (r) {
                if (r) {
                    var mms = new spa.Metadata();
                    handeler.loader("div#tb-data-toolbar-data", "loader-tb-data", "tb-data-toolbar-loader");
                    var node = $('#metaDataTree').tree('getSelected');
                    mms.deleteTermGroupbyIdAsync(getmmsid(), node.Id.toString())
                        .fail(function (e) {
                            handeler.removeloader("#tb-data-toolbar-loader");
                            sputils.flashNotificationError(handeler.parseSPException(e));
                        }).always(function () {
                            app.Default.loadMetadata();
                            handeler.removeloader("#tb-data-toolbar-loader");

                        });
                }
            });
        }

        obj.contextMenuTermSetCreateTerm = function () {
            $.messager.prompt('Insert', 'Term Name :', function (r) {
                if (r) {
                    var mms = new spa.Metadata();
                    handeler.loader("div#tb-data-toolbar-data", "loader-tb-data", "tb-data-toolbar-loader");

                    var node = $('#metaDataTree').tree('getSelected');

                    mms.createTermAsync(getmmsid(), node.Id.toString(), SP.Guid.newGuid().toString(), r, getDefailtLanguage())
                        .fail(function (e) {
                            sputils.flashNotificationError(handeler.parseSPException(e));
                        })
                        .always(function () {
                            app.Default.loadMetadata();
                            handeler.removeloader("#tb-data-toolbar-loader");
                        });
                }
            });
        }

        obj.contextMenuTermSetDeleteTermSet = function () {
            $.messager.confirm('Confirm', defs.DeleteConfirmation, function (r) {
                if (r) {
                    var mms = new spa.Metadata();
                    handeler.loader("div#tb-data-toolbar-data", "loader-tb-data", "tb-data-toolbar-loader");
                    var node = $('#metaDataTree').tree('getSelected');
                    mms.deleteTermSetAsync(getmmsid(), node.Id.toString())
                        .fail(function (e) {
                            sputils.flashNotificationError(handeler.parseSPException(e));
                        }).always(function () {
                            app.Default.loadMetadata();
                            handeler.removeloader("#tb-data-toolbar-loader");
                        });
                }
            });
        }

        obj.contextMenuTermCreateTerm = function () {
            $.messager.prompt('Insert', 'Term Name :', function (r) {
                if (r) {
                    var mms = new spa.Metadata();
                    handeler.loader("div#tb-data-toolbar-data", "loader-tb-data", "tb-data-toolbar-loader");

                    var node = $('#metaDataTree').tree('getSelected');

                    mms.createChildTermAsync(getmmsid(), node.attributes.TermSetID, node.Id.toString(), SP.Guid.newGuid().toString(), r, getDefailtLanguage())
                         .fail(function (e) {
                             sputils.flashNotificationError(handeler.parseSPException(e));
                         }).always(function () {
                             app.Default.loadMetadata();
                             handeler.removeloader("#tb-data-toolbar-loader");
                         });
                }
            });
        }

        obj.contextMenuTermDeleteTerm = function () {
            $.messager.confirm('Confirm', defs.DeleteConfirmation, function (r) {
                if (r) {
                    var mms = new spa.Metadata();
                    var node = $('#metaDataTree').tree('getSelected');
                    handeler.loader("div#tb-data-toolbar-data", "loader-tb-data", "tb-data-toolbar-loader");
                    mms.deleteTermByIdAsync(getmmsid(), node.attributes.GroupID, node.attributes.TermSetID, node.Id.toString())
                        .fail(function (e) {
                            sputils.flashNotificationError(handeler.parseSPException(e));
                        }).always(function () {
                            app.Default.loadMetadata();
                            handeler.removeloader("#tb-data-toolbar-loader");
                        });
                }
            });
        }

        function clearSubData() {
            $('#termLanguageData').propertygrid({ data: [] });
            $('#termMetaData').propertygrid({ data: [] });
        }

        function saveOtherLable() {

            handeler.loader("div#tb-data-toolbar-data", "loader-tb-data", "tb-data-toolbar-loader");
            var mms = new spa.Metadata();

            var mmsid = getmmsid();
            var node = $('#metaDataTree').tree('getSelected');
            var lcid = $("select#slcttermlanguages option:selected").val();

            var val = $("input[name='dlg-add-Labels-label']").val();

            if ($.trim(val)) {
                mms.createTermLableAsync(mmsid, node.attributes.GroupID, node.attributes.TermSetID, node.Id, lcid, val, false)
                    .fail(function (e) {
                        sputils.flashNotificationError(handeler.parseSPException(e));
                    }).always(function () {
                        app.Default.mmstermLanguageData();
                        handeler.removeloader("#tb-data-toolbar-loader");
                        $('div#dlg-add-Labels').dialog('close');
                    });
            }
            else {
                sputils.flashNotificationInfo(defs.Info_ValueCannotbeEmpty);
            }



        };

        function saveCustomProperties() {

            handeler.loader("div#tb-data-toolbar-data", "loader-tb-data", "tb-data-toolbar-loader");
            var mms = new spa.Metadata();

            var mmsid = getmmsid();
            var node = $('#metaDataTree').tree('getSelected');

            var vn = $("input[name='dlg-add-Properties-name']").val();
            var vv = $("input[name='dlg-add-Properties-value']").val();
            var vt = $("input[name='dlg-add-Properties-type']:checked").val()

            if ($.trim(vn) && $.trim(vv)) {
                if (vt == "true") {
                    mms.createSharedCustomPropertyByAsync(mmsid, node.attributes.GroupID, node.attributes.TermSetID, node.Id, vn, vv)
                        .fail(function (e) {
                            sputils.flashNotificationError(handeler.parseSPException(e));
                        }).always(function () {
                            app.Default.mmstermLanguageData();
                            handeler.removeloader("#tb-data-toolbar-loader");
                            $('div#dlg-add-Properties').dialog('close');
                        });
                }
                else {

                    mms.createLocalCustomPropertyByAsync(mmsid, node.attributes.GroupID, node.attributes.TermSetID, node.Id, vn, vv)
                         .fail(function (e) {
                             sputils.flashNotificationError(handeler.parseSPException(e));
                         })
                        .always(function () {
                            app.Default.mmstermLanguageData();
                            handeler.removeloader("#tb-data-toolbar-loader");
                            $('div#dlg-add-Properties').dialog('close');
                        });
                }
            }
            else {
                sputils.flashNotificationInfo(defs.Info_ValueCannotbeEmpty);
            }


        };

        // polutate values form Group object
        //dl = Default Language
        function populateImportGroup(group, objGrp, dl) {
            group.set_description(objGrp.d);
            return group;
        };

        // polutate values form TermSet object
        //dl = Default Language
        function populateImportTermSet(termset, objTS, dl) {
            //TODO: implementation need to be done
            return termset;
        };

        // polutate values form Term object
        //dl = Default Language
        function populateImportTerm(term, objT, dl) {

            term.setDescription(objT.d, dl);
            term.set_isAvailableForTagging(objT.a);
            // lables
            $.each(objT.z, function (k, Z) {
                //TODO : handle default
                term.createLabel(Z.v, Z.c, true);
            });
            // shared properties
            $.each(objT.y, function (l, Y) {
                term.setCustomProperty(Y.n, Y.v);
            });
            // custom properties
            $.each(objT.w, function (m, W) {
                term.setLocalCustomProperty(W.n, W.v);
            });

            return term;
        };

        // create child term in parent term as a loop
        // PT=Parent Term , rt=reference Term , dl=Default Language , options as optional
        function createChildTerms(PT, rt, dl, options) {

            $.each(rt.t, function (i, ct) {

                // override the guide for new instance
                if (options && options[enums.createChildTermsOptions.NewGuid]) {
                    ct.i = SP.Guid.newGuid().toString();
                }

                var cterm = PT.createTerm(ct.x, dl, ct.i);
                cterm = populateImportTerm(cterm, ct, dl);

                if (ct.t.length > 0) {
                    createChildTerms(cterm, ct, dl, options);
                }
            });
        };

        obj.saveChangestermProperties = function () {
            var rows = $('#termMetaData').propertygrid('getChanges');

            var mmsid = getmmsid();
            var node = $('#metaDataTree').tree('getSelected');

            var mms = new spa.Metadata();

            var lcid = $("select#slcttermlanguages").attr("defaultl");

            var calls = [];
            handeler.loader("div#tb-data-toolbar-data", "loader-tb-data", "tb-data-toolbar-loader");

            for (var i = 0; i < rows.length; i++) {
                switch (rows[i].name) {
                    case "Description":
                        calls.push(mms.setTermDescriptionById(mmsid, node.attributes.GroupID, node.attributes.TermSetID, node.Id, lcid, rows[i].value));
                        break;
                    case "Tagging":
                        calls.push(mms.setTermTaggingById(mmsid, node.attributes.GroupID, node.attributes.TermSetID, node.Id, rows[i].value));
                        break;
                    case "Name":

                        var val = rows[i].value;
                        var cName = mms.setTermNameById(mmsid, node.attributes.GroupID, node.attributes.TermSetID, node.Id, lcid, val);
                        cName.done(function () {
                            var node = $('#metaDataTree').tree('getSelected');
                            if (node) {
                                $('#metaDataTree').tree('update', {
                                    target: node.target,
                                    text: val
                                });
                            }
                        });

                        break;
                    default:

                }
            }
            $.when.apply($, calls)
                .fail(function (e) {
                    sputils.flashNotificationError(handeler.parseSPException(e));
                })
                .done(function () {
                    sputils.flashNotificationInfo(defs.Info_Saved);
                })
                .always(function () {
                    handeler.removeloader("#tb-data-toolbar-loader");
                });
        }

        obj.selectmmsItem = function (current) {

            handeler.waitUIProcess();

            var c = $(current);
            var area = $("span#DeltaPlaceHolderPageTitleInTitleArea");

            area.text(c.attr("data-name"));
            area.attr("data-mmsid", c.attr("data-mmsid"));
            area.attr("data-dl", c.attr("data-dl"));
            area.attr("data-l", c.attr("data-l"));

            app.Default.onPropertyLoad();
        };

        //load term stores to select
        obj.onLoad = function () {

            controlInt();

            appInstalation(
                function () {

                    $$.IncludeScript($$.js.spTaxonomy, function () {

                        var mms = new spa.Metadata();

                        mms.getTermStoresAsync().done(function (mms) {
                            var enu = mms.getEnumerator();
                            while (enu.moveNext()) {
                                generatemmsitem(enu.get_current());
                            }
                        }).fail(function (e) {
                            sputils.flashNotificationError(handeler.parseSPException(e));
                        }).always(function () { handeler.doneWaiting(); });
                    });
                },
                function () {
                    handeler.doneWaiting();
                    sputils.flashNotificationError(defs.Error_AppInitialization);
                }
                );
        };

        // Load the metada tree in the first time in the data tab
        obj.loadMetadata = function () {
            $$.IncludeScript($$.js.spTaxonomy, function () {

                handeler.waitUIElement("div#TreeCategories");

                var mmsid = getmmsid();
                var mms = new spa.Metadata();

                mms.getTermGroupsbyStoreIdAsync(mmsid)
                    .done(function (itemgrps) {

                        try {
                            var enu = itemgrps.getEnumerator();
                            var treeData = [];
                            while (enu.moveNext()) {
                                var cu = enu.get_current();
                                treeData.push(groupToTreeNode(cu));
                            }

                            var rootData = [];
                            var root = new mapers.TreeItem();
                            root.Id = mmsid;
                            root.text = "Managed MetaData";
                            root.state = mapers.TreeItem.State.Open;
                            root.children = treeData;
                            root.attributes = { "Store": true };
                            rootData.push(root);

                            $('#metaDataTree').tree(
                                {
                                    data: rootData,
                                    checkbox: false,
                                    onContextMenu: function (e, node) {
                                        e.preventDefault();

                                        $(this).tree('select', node.target);

                                        if (node.attributes.Store) {
                                            $('#metaDataTree-context-store').menu('show', {
                                                left: e.pageX,
                                                top: e.pageY
                                            });
                                        }
                                        else if (node.attributes.Group) {
                                            $('#metaDataTree-context-group').menu('show', {
                                                left: e.pageX,
                                                top: e.pageY
                                            });
                                        }
                                        else if (node.attributes.TermSet) {
                                            $('#metaDataTree-context-termset').menu('show', {
                                                left: e.pageX,
                                                top: e.pageY
                                            });
                                        }
                                        else if (node.attributes.Term) {
                                            $('#metaDataTree-context-term').menu('show', {
                                                left: e.pageX,
                                                top: e.pageY
                                            });
                                        }
                                    }
                                });

                        } catch (ex) {
                            sputils.flashNotificationError(defs.Error_Reshesh);
                        }
                    }).fail(function (e) {
                        $.messager.alert(defs.Msgs.Error_Title, handeler.parseSPException(e), 'error');
                    }).always(function () {
                        disabletabs(false); // First call
                        clearSubData();
                        handeler.doneWaiting();
                        handeler.donewaitElement("div#TreeCategories");
                    });
            });
        };

        // Load the metada tree template in the first time in the template tab
        obj.loadMetadataTemplete = function () {
            $$.IncludeScript($$.js.spTaxonomy, function () {

                handeler.waitUIElement("div#TemplateTree");

                var mmsid = getmmsid();
                var mms = new spa.Metadata();

                mms.getTermGroupsbyStoreIdAsync(mmsid)
                    .done(function (itemgrps) {

                        try {
                            var enu = itemgrps.getEnumerator();
                            var treeData = [];
                            while (enu.moveNext()) {
                                var cu = enu.get_current();
                                treeData.push(groupToTreeNode(cu));
                            }

                            var rootData = [];
                            var root = new mapers.TreeItem();
                            root.Id = mmsid;
                            root.text = "Managed MetaData";
                            root.state = mapers.TreeItem.State.Open;
                            root.children = treeData;
                            root.attributes = { "Store": true };
                            rootData.push(root);

                            $('#templateTree').tree({
                                data: rootData,
                                checkbox: false,
                                onContextMenu: function (e, node) {
                                    e.preventDefault();
                                    $(this).tree('select', node.target);

                                    if (node.attributes[mapers.TreeItem.Attributes.isTermStore]) {
                                        $('#templateTree-context-root').menu('show', { left: e.pageX, top: e.pageY });
                                    }
                                    else {
                                        $('#templateTree-context').menu('show', { left: e.pageX, top: e.pageY });
                                    }
                                }
                            });

                        } catch (ex) {
                            sputils.flashNotificationError(defs.Error_Reshesh);
                        }
                    }).fail(function (e) {
                        $.messager.alert(defs.Msgs.Error_Title, handeler.parseSPException(e), 'error');
                    }).always(function () {
                        handeler.doneWaiting();
                        handeler.donewaitElement("div#TemplateTree");
                    });
            });
        };

        // load child data when tree node is clicked 
        // tree is used in two tabs so tree selector is passed to the method
        obj.loadChildData = function (node, treeSelector) {

            var mms = new spa.Metadata();
            var mmsid = getmmsid();
            var ret = false;

            if (node.children.length == 0 && node.state == mapers.TreeItem.State.Close) {

                node.state == mapers.TreeItem.State.Open; // Make node open

                if (node.attributes.Group) {
                    ret = false;
                    handeler.waitUIProcess();

                    mms.getallTermSetsbyGroupAsync(mmsid, node.Id)
                        .done(function (termsets) {

                            var enu = termsets.getEnumerator();
                            treeData = [];
                            while (enu.moveNext()) {
                                var cu = enu.get_current();
                                treeData.push(termSetToTreeNode(cu, node.Id));
                            }

                            if (treeData.length == 0) {
                                node.state == mapers.TreeItem.State.Open;
                            }


                            $(treeSelector).tree('append', {
                                parent: node.target,
                                data: treeData
                            }).tree('expandTo', node.target);
                        })
                        .fail(function (e) {
                            sputils.flashNotificationError(handeler.parseSPException(e));
                        })
                        .always(function () {
                            handeler.doneWaiting();
                        });
                }
                else if (node.attributes.TermSet) {
                    ret = false;
                    handeler.waitUIProcess();

                    mms.getTermsAsync(mmsid, node.attributes.GroupID, node.Id)
                        .done(function (terms) {

                            var enu = terms.getEnumerator();
                            treeData = [];
                            while (enu.moveNext()) {
                                var cu = enu.get_current();
                                treeData.push(termToTreeNode(cu, node.attributes.GroupID, node.Id));
                            }

                            if (treeData.length == 0) {
                                node.state == mapers.TreeItem.State.Open;
                            }

                            $(treeSelector).tree('append', {
                                parent: node.target,
                                data: treeData
                            }).tree('expandTo', node.target);
                        }).fail(function (e) {
                            sputils.flashNotificationError(handeler.parseSPException(e));
                        })
                        .always(function () {
                            handeler.doneWaiting();
                        });
                }
                else if (node.attributes.Term) {
                    ret = false;
                    handeler.waitUIProcess();

                    mms.getChildTermsAsync(mmsid, node.attributes.GroupID, node.attributes.TermSetID, node.Id)
                        .done(function (items) {
                            treeData = [];
                            var enu = items.getEnumerator();
                            while (enu.moveNext()) {
                                var cu = enu.get_current();
                                treeData.push(termToTreeNode(cu, node.attributes.GroupID, node.attributes.TermSetID));
                            }

                            if (treeData.length == 0) {
                                node.state == mapers.TreeItem.State.Open;
                            }

                            $(treeSelector).tree('append', {
                                parent: node.target,
                                data: treeData
                            }).tree('expandTo', node.target);
                        }).fail(function (e) {
                            sputils.flashNotificationError(handeler.parseSPException(e));
                        })
                        .always(function () {
                            handeler.doneWaiting();
                        });
                };
            } else {
                ret = true;
            };

            return ret;

        };

        // load term store properites and load the trees in other tabs
        obj.onPropertyLoad = function () {

            var mmsid = getmmsid();
            var mms = new spa.Metadata();

            mms.getTermStoreByIdAsync(mmsid)
                .done(function (mmsstore) {
                    handeler.waitUIElement("div#tbproperties");
                    var datal = new mapers.PropertyData();
                    datal.addRow("Name", mmsstore.get_name(), "General", "");
                    datal.addRow("Online", mmsstore.get_isOnline(), "General", "");
                    datal.addRow("Default ", enums.Languages[mmsstore.get_defaultLanguage()], "Language", "");
                    datal.addRow("Working", enums.Languages[mmsstore.get_workingLanguage()], "Language", "");

                    var tmplan = "";
                    $("select#slcttermlanguages").empty();
                    $(mmsstore.get_languages()).each(
                        function (a, b) {
                            $("select#slcttermlanguages").append("<option value='" + b + "'>" + enums.Languages[b] + "</option>");
                            tmplan += enums.Languages[b] + "; ";
                        });

                    $("select#slcttermlanguages").val(mmsstore.get_defaultLanguage());
                    $("select#slcttermlanguages").attr("defaultl", mmsstore.get_defaultLanguage());

                    datal.addRow("Languages", tmplan, "Language", "");

                    $('#pgMetaData').propertygrid({
                        data: datal,
                        showGroup: true,
                        scrollbarSize: 0,
                        fitColumns: false,
                        columns: [[{ field: 'name', title: 'Settings', width: 80 }, { field: 'value', title: 'Value' }]]
                    });

                    disabletabs(true);
                    app.Default.loadMetadata();
                    app.Default.loadMetadataTemplete();

                }).fail(function (e) {
                    sputils.flashNotificationError(handeler.parseSPException(e));
                })
        .always(function () {
            handeler.donewaitElement("div#tbproperties");
        });


        };

        // Load general information node click meta data tree
        obj.mmsTreeNodeClick = function (node) {
            var mms = new spa.Metadata();
            if (node.attributes.Term) {
                var mmsid = getmmsid();
                mms.getTermByIdAsync(mmsid, node.attributes.GroupID, node.attributes.TermSetID, node.Id)
                    .done(function (term) {

                        var datal = new mapers.PropertyData();
                        var tmp;

                        datal.addRow("Name", term.get_name(), "General Properties (Editable)", "text");
                        datal.addRow("Description", term.get_description(), "General Properties (Editable)", "text");
                        datal.addRow("Tagging", term.get_isAvailableForTagging(), "General Properties (Editable)", { "type": "checkbox", "options": { "on": true, "off": false } });

                        tmp = new Date(term.get_createdDate());
                        datal.addRow("Created", tmp.format("MMM dd yyyy HH:mm:ss 'GMT'zzz"), "System", "");
                        tmp = new Date(term.get_lastModifiedDate());
                        datal.addRow("Modified", tmp.format("MMM dd yyyy HH:mm:ss 'GMT'zzz"), "System", "");

                        datal.addRow("Deprecated", term.get_isDeprecated(), "System", "");
                        datal.addRow("Keyword", term.get_isKeyword(), "System", "");
                        datal.addRow("Pinned", term.get_isPinned(), "System", "");
                        datal.addRow("Pinned Root", term.get_isPinnedRoot(), "System", "");
                        datal.addRow("Reused", term.get_isReused(), "System", "");
                        datal.addRow("Root", term.get_isRoot(), "System", "");
                        datal.addRow("Source Term", term.get_isSourceTerm(), "System", "");

                        $('#termMetaData').propertygrid({
                            data: datal,
                            showGroup: true,
                            showHeader: false,
                            scrollbarSize: 0,
                            fitColumns: true
                        });

                        $("select#slcttermlanguages").val($("select#slcttermlanguages").attr("defaultl"));
                        app.Default.mmstermLanguageData();
                    })
                    .fail(function (e) {
                        sputils.flashNotificationError(handeler.parseSPException(e));
                    });
            }
            else {
                clearSubData(); // Clear tables
            }

        };

        // loado language data for selected term
        obj.mmstermLanguageData = function () {
            var mms = new spa.Metadata();

            var mmsid = getmmsid();

            var nd = $('#metaDataTree').tree('getSelected')

            if (nd) {

                var datal = new mapers.PropertyData();

                mms.getTermByIdAsyncWithLabels(mmsid, nd.attributes.GroupID, nd.attributes.TermSetID, nd.Id, $("select#slcttermlanguages option:selected").val())
                    .done(function (lables) {
                        var enu = lables.getEnumerator();
                        while (enu.moveNext()) {
                            var c = enu.get_current();

                            if (c.get_isDefaultForLanguage()) {
                                datal.addRow("Synonyms & Abbreviations", c.get_value(), "Default Label", "");
                            }
                            else {
                                datal.addRow("Synonyms & Abbreviations", c.get_value(), "Other Labels", "");
                            }
                        }

                        mms.getTermSharedPropertiesByAsync(mmsid, nd.attributes.GroupID, nd.attributes.TermSetID, nd.Id)
                            .done(function (properties) {
                                $.each(properties, function (k, v) {
                                    datal.addRow(k, v, "Shared Properties", "");
                                });
                                mms.getTermLocalPropertiesByAsync(mmsid, nd.attributes.GroupID, nd.attributes.TermSetID, nd.Id)
                                    .done(function (properties) {
                                        $.each(properties, function (k, v) {
                                            datal.addRow(k, v, "Local Properties", "");
                                        });
                                        $('#termLanguageData').propertygrid({
                                            data: datal,
                                            showGroup: true,
                                            showHeader: false,
                                            scrollbarSize: 0,
                                            fitColumns: true
                                        });
                                    }).fail(function (e) {
                                        sputils.flashNotificationError(handeler.parseSPException(e));
                                    });
                            }).fail(function (e) {
                                sputils.flashNotificationError(handeler.parseSPException(e));
                            });
                    }).fail(function (e) {
                        sputils.flashNotificationError(handeler.parseSPException(e));
                    });
            }
        };

        // open dlg for add lable
        obj.addLabels = function () {
            var node = $('#metaDataTree').tree('getSelected');
            if (node && node.attributes.Term) {
                $('div#dlg-add-Labels').dialog('open');
            }
            else {
                sputils.flashNotificationInfo(defs.Info_SelectTerm);
            }
        };

        // open dlg for add CustomProperty
        obj.addCustomProperty = function () {
            var node = $('#metaDataTree').tree('getSelected');
            if (node && node.attributes.Term) {
                $('div#dlg-add-Properties').dialog('open');
            }
            else {
                sputils.flashNotificationInfo(defs.Info_SelectTerm);
            }
        };

        // delete language data for selected term
        obj.deleteLanguageData = function () {
            var selected = $('#termLanguageData').propertygrid('getSelected');
            var node = $('#metaDataTree').tree('getSelected');

            if (selected && node && node.attributes.Term) {

                $.messager.confirm('Confirm', defs.DeleteConfirmation, function (r) {
                    if (r) {
                        handeler.loader("div#tb-data-toolbar-data", "loader-tb-data", "tb-data-toolbar-loader");

                        var mms = new spa.Metadata();

                        var mmsid = getmmsid();
                        var lcid = $("select#slcttermlanguages option:selected").val();

                        switch (selected.group) {
                            case "Shared Properties":
                                mms.deleteSharedCustomPropertyByAsync(mmsid, node.attributes.GroupID, node.attributes.TermSetID, node.Id, selected.name)
                                    .fail(function (e) {
                                        sputils.flashNotificationError(handeler.parseSPException(e));
                                    }).always(function () {
                                        app.Default.mmstermLanguageData();
                                        handeler.removeloader("#tb-data-toolbar-loader");
                                    });
                                break;
                            case "Local Properties":
                                mms.deleteLocalCustomPropertyByAsync(mmsid, node.attributes.GroupID, node.attributes.TermSetID, node.Id, selected.name)
                                    .fail(function (e) {
                                        sputils.flashNotificationError(handeler.parseSPException(e));
                                    }).always(function () {
                                        app.Default.mmstermLanguageData();
                                        handeler.removeloader("#tb-data-toolbar-loader");
                                    });
                                break;
                            case "Other Labels":
                                mms.deletelableByAsync(mmsid, node.attributes.GroupID, node.attributes.TermSetID, node.Id, lcid, selected.value)
                                    .fail(function (e) {
                                        sputils.flashNotificationError(handeler.parseSPException(e));
                                    }).always(function () {
                                        app.Default.mmstermLanguageData();
                                        handeler.removeloader("#tb-data-toolbar-loader");
                                    });
                                break;
                            default:

                        }
                    }
                });
            }
            else {
                sputils.flashNotificationInfo(defs.Info_SelectItem);
            }

        };

        obj.TabImportExport = (function () {

            var obj = {};
            var objdata = ei.getObjectData();

            // object used for Async call templates
            function objdataTemplate() {
                objdata = ei.getObjectData();
                objdata.tm.v = "ex.1.0.0.0";
                objdata.tm.dl = getDefailtLanguage(); // get term store default language
                objdata.tm.l = getMMSLanguages();
            };

            function callback() {

                setTimeout(function () {

                    if (objdata.rc == objdata.cc) {

                        var op = new spa.Operations();

                        var itm = op.getNewListItem(defs.L.Export.Nm);

                        var data = B64.encode(JSON.stringify(objdata.tm));

                        itm.set_item(defs.L.Export.C.ExportData, data);
                        itm.set_item(defs.L.Export.C.ExportLog, objdata.lg);
                        itm.set_item(defs.L.Export.C.Title, objdata.tm.Name);

                        itm.update();

                        op.insertListItem(itm).done(function () {
                            exportSuccess();
                        }).fail(function () {
                            exportCanceled();;
                        });

                    }
                    else if (objdata.fc > 0) {
                        exportCanceled();
                    }
                    else {
                        callback();
                    }
                }, 2000);
            }

            function exportCanceled() {
                $.messager.progress('close');
                $.messager.alert(defs.String.Dlg.Title_Error, defs.String.Err.OperationFailed, 'error');
                objdataTemplate();
            }

            function exportSuccess() {
                $.messager.progress('close');
                $.messager.alert(defs.String.Dlg.Title_General, defs.String.Info.ExportSaveSuccess, 'info');
                objdataTemplate();
                obj.loadExportGrid(); // load the grid after saving
            }

            function importSuccess() {
                $.messager.progress('close');
                $.messager.alert('Import', defs.ImportOperationSuccess, 'info');

                // clear the object
                objdataTemplate();
            }

            function importFailed() {
                $.messager.progress('close');
                $.messager.alert('Import', defs.OperationFailed, 'error');

                //
                objdataTemplate();
            };

            function importCallBack() {
                setTimeout(function () {

                    if (objdata.rc == objdata.cc) {
                        importSuccess();
                    }
                    else if (objdata.fc > 0) {

                        //TODO; show the log
                        importFailed();
                    }
                    else {
                        importCallBack();
                    }
                }, 2000);
            };

            // tab load
            obj.loadTab = function () {
                obj.loadExportGrid();
            };

            // load export data to the grid
            obj.loadExportGrid = function () {

                var loaderId = "tb-data-toolbar-loader";
                handeler.loader("div#tb-exportimport-toolbar-export", "loader-tb-data", loaderId);
                spREST.Web.Lists.get(defs.L.Export.Nm, defs.L.Export.Q.getAllExports).done(function (data) {

                    var dataSet = [];
                    $.each(data.d.results, function (i, ele) {
                        var x = new Date(parseInt(ele[defs.L.Export.C.Created].substring(6).substring(0, 13)));
                        dataSet.push({ Id: ele[defs.L.Export.C.Id], tmp: ele[defs.L.Export.C.Title], dt: x.toGMTString() });
                    });

                    loadDataToExportGrid(dataSet);

                }).fail(function (e) {
                    sputils.flashNotificationError(handeler.parseSPException(e));
                }).always(function () {
                    handeler.removeloader("#" + loaderId);
                });

            };

            // export term store
            obj.exportTermStore = function () {

                objdataTemplate();
                objdata.Options.Export.TermLables = true;
                objdata.Options.Export.TermProperties = true;


                var mms = new spa.Metadata();
                $.messager.prompt(defs.String.Dlg.Title_General, defs.String.Dlg.EnterName, function (r) {
                    if (r) {

                        objdata.tm.Type = mapers.Template.Type.TermStore;
                        objdata.tm.Name = r;

                        var TermStore = new mapers.TS();
                        TermStore.i = getmmsid();
                        TermStore.n = getmmsname();
                        TermStore.dl = getDefailtLanguage();
                        TermStore.l = getMMSLanguages();
                        objdata.tm.Data = TermStore;

                        $.messager.progress({ msg: defs.String.Info.DoNotCloseBrowser, text: defs.String.Info.Exportinprogress });

                        mms.getTermGroupsbyStoreIdAsync(objdata.tm.Data.i).done(function (grps) {
                            var enu = grps.getEnumerator();
                            while (enu.moveNext()) {
                                var cu = enu.get_current();
                                var g = ei.mapGroupObject(cu);

                                objdata.tm.Data.g.push(g);
                            }

                            ei.exportGetTermSetsandTerms(objdata);
                            //asyc the calls
                            callback();

                        }).fail(function () {
                            // close the msg
                            objdata.fc++;
                            callback();
                        });

                    }
                });

            };

            // click export grid load preview
            obj.clickExportGrid = function (rowIndex, rowData) {

                $("ul#exportTreePreview").empty();

                var loaderId = "tb-data-toolbar-loader";
                handeler.loader("div#tb-exportimport-toolbar-export", "loader-tb-data", loaderId);
                spREST.Web.Lists.get(defs.L.Export.Nm, defs.L.Export.Q.getExportById(rowData.Id)).done(function (data) {

                    var obj = JSON.parse(B64.decode(data.d.results[0][defs.L.Export.C.ExportData]));
                    loadExportMetaDataToTree("ul#exportTreePreview", obj.Data)

                }).fail(function (e) {
                    sputils.flashNotificationError(handeler.parseSPException(e));
                }).always(function () {
                    handeler.removeloader("#" + loaderId);
                });

            };

            obj.openImportDialog = function () {

                var op = new spa.Operations();

                // get the selected row for id
                var row = $('#exportGrid').datagrid('getSelected');

                if (row) {


                    var loaderId = "tb-data-toolbar-loader";
                    handeler.loader("div#tb-exportimport-toolbar-export", "loader-tb-data", loaderId);
                    spREST.Web.Lists.get(defs.L.Export.Nm, defs.L.Export.Q.getExportById(row.Id)).done(function (data) {

                        var obj = JSON.parse(B64.decode(data.d.results[0][defs.L.Export.C.ExportData]));
                        //loadExportMetaDataToTree("ul#exportTreePreview", obj.Data)

                        var treeData = [];

                        obj.Data.g.forEach(function (G) {
                            var itmg = itemToNodeExportTree(G);
                            itmg.attributes[mapers.TreeItem.Attributes.isGroup] = true; // set group attibute
                            treeData.push(itmg);
                        });

                        var rootData = [];
                        var root = new mapers.TreeItem();
                        root.Id = obj.Data.i;
                        root.text = "Import : " + obj.Name;
                        root.state = mapers.TreeItem.State.Close;
                        root.children = treeData;
                        rootData.push(root);

                        $('ul#ImporttermaDataTree').tree({ data: rootData, checkbox: true });

                        $('div#dlg-import-terms').dialog('open');


                    }).fail(function (e) {
                        sputils.flashNotificationError(handeler.parseSPException(e));
                    }).always(function () {
                        handeler.removeloader("#" + loaderId);
                    });

                }
                else {
                    $.messager.alert('Import', defs.ImportPleaseSelect, 'info');
                };

            };

            function importaData(dataToImport) {

                var md = new spa.Metadata();

                $('div#dlg-import-terms').dialog('close');
                $.messager.progress({ msg: "Do not close the browser or refresh ", text: "Import in progress" });

                $.each(dataToImport.Data.g, function (i, Grp) {

                    var grp = md.getCreatedGroup(getmmsid(), Grp.i, Grp.x);
                    grp = populateImportGroup(grp, dataToImport.Data.g[i], dataToImport.dl);

                    $.each(dataToImport.Data.g[i].s, function (ii, TS) {
                        var ts = grp.createTermSet(TS.x, TS.i, dataToImport.dl);
                        ts = populateImportTermSet(ts, dataToImport.Data.g[i].s[ii], dataToImport.dl);

                        // loop each term
                        $.each(TS.t, function (j, T) {
                            var t = ts.createTerm(T.x, dataToImport.dl, T.i);
                            t = populateImportTerm(t, dataToImport.Data.g[i].s[ii].t[j], dataToImport.dl);

                            if (T.t.length > 0) {
                                createChildTerms(t, T, dataToImport.dl);
                            };
                        });
                    });

                    dataToImport.cc++;
                    spa.executeCtxAsync(grp).done(function () {
                        dataToImport.rc++;
                    }).fail(function (b) {
                        dataToImport.fc++;
                        dataToImport.log(dataToImport.Data.g[i].x + ":" + b);
                    });

                });

                importCallBack();
            }

            obj.importExportedData = function () {

                var vt = $("input[name='dlg-import-terms-options']:checked").val()
                var checkedg = $('ul#ImporttermaDataTree').tree('getChecked');

                if (checkedg) {
                    var mms = new spa.Metadata();

                    mms.getTermGroupsbyStoreIdAsync(getmmsid()).done(function (grops) {
                        var enu = grops.getEnumerator();
                        var notExist = true;

                        var serverGrps = [];

                        while (enu.moveNext()) {
                            var c = enu.get_current();
                            serverGrps.push(c.get_id().toString());
                        };

                        objdataTemplate(); // init object data

                        // load the object data
                        var loaderId = "tb-data-toolbar-loader";
                        handeler.loader("div#tb-exportimport-toolbar-export", "loader-tb-data", loaderId);
                        var rowData = $('#exportGrid').datagrid('getSelected');
                        spREST.Web.Lists.get(defs.L.Export.Nm, defs.L.Export.Q.getExportById(rowData.Id)).done(function (data) {

                            var objTemplate = JSON.parse(B64.decode(data.d.results[0][defs.L.Export.C.ExportData]));

                            switch (vt) {
                                case "1":
                                    break;
                                case "2": // Fail if exist
                                    var Exist = checkedg.some(function (ele, index, array) {
                                        return serverGrps.some(function (elemesg, i, ar) {
                                            return elemesg == ele.Id
                                        });
                                    });
                                    if (Exist) {
                                        $('div#dlg-import-terms').dialog('close');
                                        $.messager.alert('Import', "One or more Groups exist in the selected term store. Import canceled", 'warning');
                                    }
                                    else {

                                        objTemplate.Data.g = objTemplate.Data.g.filter(function (element) {
                                            return checkedg.some(function (ele, index, array) { return ele.Id == element.i; });
                                        });

                                        importaData(vt, objTemplate);
                                    };

                                    break;
                                case "3": // Skip if exist

                                    var toimport = [];

                                    $.each(checkedg, function (i, v) {

                                        if (v.attributes.Group) {
                                            var notexist = !serverGrps.some(function (elemesg, i, ar) {
                                                return elemesg == v.Id
                                            });

                                            if (notexist) {
                                                toimport.push(v.Id);
                                            }
                                        }
                                    });

                                    objTemplate.Data.g = objTemplate.Data.g.filter(function (element) {
                                        return toimport.some(function (ele, index, array) { return ele == element.i; });
                                    });

                                    importaData(objTemplate);

                                    break;
                                default:
                                    break;
                            }


                        }).fail(function (e) {
                            sputils.flashNotificationError(handeler.parseSPException(e));
                        }).always(function () {
                            handeler.removeloader("#" + loaderId);
                        });

                    }).fail(function (b) {
                        $('div#dlg-import-terms').dialog('close');
                        sputils.flashNotificationError(handeler.parseSPException(b));
                    });
                };

            };

            return obj;
        })();

        obj.TabTemplate = (function () {

            var obj = {};

            //var objdata = { "tm": new mapers.Template(), "cc": 0, "rc": 0, "fc": 0, "lg": "" };
            var objdata = ei.getObjectData();

            // object used for Async call templates
            function objdataTemplate() {
                objdata = ei.getObjectData();
                objdata.tm.v = "tm.1.0.0.0";
                objdata.tm.dl = getDefailtLanguage(); // get term store default language
                objdata.tm.l = getMMSLanguages();
                objdata.Options.Export.TermLables = true;
                objdata.Options.Export.TermProperties = true;

            };

            function callback() {

                setTimeout(function () {

                    if (objdata.rc == objdata.cc) {

                        var op = new spa.Operations();

                        var itm = op.getNewListItem(defs.L.Templates.Nm);

                        var data = B64.encode(JSON.stringify(objdata.tm));

                        itm.set_item(defs.L.Templates.C.TemplateType, objdata.tm.Type);
                        itm.set_item(defs.L.Templates.C.TemplateData, data);
                        itm.set_item(defs.L.Templates.C.TemplateLog, objdata.lg);
                        itm.set_item(defs.L.Templates.C.Title, objdata.tm.Name);

                        itm.update();

                        op.insertListItem(itm).done(function () {
                            templateSuccess();
                        }).fail(function () {
                            templateCanceled();;
                        });

                    }
                    else if (objdata.fc > 0) {
                        templateCanceled();
                    }
                    else {
                        callback();
                    }
                }, 2000);
            }

            function templateCanceled() {
                $.messager.progress('close');
                $.messager.alert(defs.String.Dlg.Title_Error, defs.OperationFailed, 'error');
                objdataTemplate();
            }

            function templateSuccess() {
                $.messager.progress('close');
                $.messager.alert(defs.String.Dlg.Title_General, defs.String.Info.TemplateSaveSuccess, 'info');
                objdataTemplate();
                obj.LoadTemplates(); // load the grid after saving
            }

            // Save as Term Store Template
            function exportAsTermStoreTemplate(node, funcallback) {
                var mms = new spa.Metadata();
                $.messager.prompt(defs.String.Dlg.Title_General, defs.String.Dlg.EnterTemplateName, function (r) {
                    if (r) {

                        objdata.tm.Type = mapers.Template.Type.TermStore;
                        objdata.tm.Name = r;

                        var TermStore = new mapers.TS();
                        TermStore.i = getmmsid();
                        TermStore.n = node.text;
                        TermStore.dl = getDefailtLanguage();
                        TermStore.l = getMMSLanguages();
                        objdata.tm.Data = TermStore;

                        $.messager.progress({ msg: defs.String.Info.DoNotCloseBrowser, text: defs.String.Info.Exportinprogress });

                        mms.getTermGroupsbyStoreIdAsync(objdata.tm.Data.i).done(function (grps) {
                            var enu = grps.getEnumerator();
                            while (enu.moveNext()) {
                                var cu = enu.get_current();
                                var g = ei.mapGroupObject(cu);

                                objdata.tm.Data.g.push(g);
                            }

                            ei.exportGetTermSetsandTerms(objdata);
                            //asyc the calls
                            funcallback();

                        }).fail(function () {
                            // close the msg
                            objdata.fc++;
                            funcallback();
                        });

                    }
                });
            };

            // Save as Group Template
            function ExportAsGroupTemplate(node, funcallback) {
                var mms = new spa.Metadata();
                $.messager.prompt(defs.String.Dlg.Title_General, defs.String.Dlg.EnterTemplateName, function (r) {
                    if (r) {

                        // Blank term Store with Id
                        var TermStore = new mapers.TS();
                        TermStore.i = getmmsid();
                        TermStore.n = "";
                        TermStore.dl = getDefailtLanguage();
                        TermStore.l = getMMSLanguages();

                        // assign selected node Id as a group id
                        var grp = new mapers.G();
                        grp.i = node.Id.toString();
                        grp.x = node.text;

                        TermStore.g.push(grp);

                        // assign objdata
                        objdata.tm.Type = mapers.Template.Type.Group;
                        objdata.tm.Name = r;
                        objdata.tm.Data = TermStore;

                        $.messager.progress({ msg: defs.String.Info.DoNotCloseBrowser, text: defs.String.Info.Exportinprogress });
                        mms.getTermGroupbyIdAsync(objdata.tm.Data.i, objdata.tm.Data.g[0].i).done(function (grp) {

                            // assign the group 
                            objdata.tm.Data.g[0] = ei.mapGroupObject(grp);
                            ei.exportGetTermSetsandTerms(objdata);

                            // make call sync
                            funcallback();

                        }).fail(function () {
                            // close the msg
                            objdata.fc++;
                            funcallback();
                        });

                    }
                });
            };

            // Save as Term Set Template
            function ExportAsTermSetTemplate(node, funcallback) {
                var mms = new spa.Metadata();
                $.messager.prompt(defs.String.Dlg.Title_General, defs.String.Dlg.EnterTemplateName, function (r) {
                    if (r) {

                        // Blank TermStore with Id
                        var TermStore = new mapers.TS();
                        TermStore.i = getmmsid();
                        TermStore.n = "";
                        TermStore.dl = getDefailtLanguage();
                        TermStore.l = getMMSLanguages();

                        // Blank Group
                        var grp = new mapers.G();
                        grp.i = node.attributes[mapers.TreeItem.Attributes.GroupID].toString();
                        grp.x = "";

                        // term Set with ID
                        var ts = new mapers.S();
                        ts.i = node.Id.toString();
                        ts.x = node.text;
                        grp.s.push(ts);

                        TermStore.g.push(grp);

                        objdata.tm.Type = mapers.Template.Type.TermSet;
                        objdata.tm.Name = r;
                        objdata.tm.Data = TermStore;

                        $.messager.progress({ msg: defs.String.Info.DoNotCloseBrowser, text: defs.String.Info.Exportinprogress });

                        // get terms for term set
                        ei.exportGetTermsForTermSet(objdata.tm.Data.g[0], objdata.tm.Data.g[0].s[0], objdata);

                        // async the calls
                        funcallback();

                    }
                });
            }

            // Save as Term Termplate
            function ExportAsTermTemplate(node, funcallback) {
                var mms = new spa.Metadata();
                $.messager.prompt(defs.String.Dlg.Title_General, defs.String.Dlg.EnterTemplateName, function (r) {
                    if (r) {

                        // Blank TermStore with Id
                        var TermStore = new mapers.TS();
                        TermStore.i = getmmsid();
                        TermStore.n = "";
                        TermStore.dl = getDefailtLanguage();
                        TermStore.l = getMMSLanguages();

                        // Blank Group
                        var grp = new mapers.G();
                        grp.i = node.attributes[mapers.TreeItem.Attributes.GroupID].toString();
                        grp.x = "";

                        // Blank TetmSet
                        var ts = new mapers.S();
                        ts.i = node.attributes[mapers.TreeItem.Attributes.TermSetID].toString();
                        ts.x = "";

                        // asign a term
                        var t = new mapers.T();
                        t.i = node.Id.toString();
                        t.x = node.text;
                        ts.t.push(t);

                        grp.s.push(ts);
                        TermStore.g.push(grp);

                        objdata.tm.Type = mapers.Template.Type.Term;
                        objdata.tm.Name = r;
                        objdata.tm.Data = TermStore;

                        $.messager.progress({ msg: defs.String.Info.DoNotCloseBrowser, text: defs.String.Info.Exportinprogress });

                        // get terms for term
                        ei.exportGetChildTerms(objdata.tm.Data.g[0].s[0].t[0], objdata.tm.Data.g[0], objdata.tm.Data.g[0].s[0], objdata);

                        // async the calls
                        funcallback();

                    }
                });
            };

            // handle save as template
            obj.ExportAsTemplate = function () {

                objdataTemplate();
                node = $('#templateTree').tree('getSelected');

                // do not execute below if node not selected
                if (!node) {
                    $.messager.alert(defs.String.Dlg.Title_General, defs.Info_SelectItem, 'info');
                    return;
                };

                if (node.attributes[mapers.TreeItem.Attributes.isTermStore]) {
                    exportAsTermStoreTemplate(node, callback);
                }
                else if (node.attributes[mapers.TreeItem.Attributes.isGroup]) {
                    ExportAsGroupTemplate(node, callback);
                }
                else if (node.attributes[mapers.TreeItem.Attributes.isTermSet]) {
                    ExportAsTermSetTemplate(node, callback);
                } else if (node.attributes[mapers.TreeItem.Attributes.isTerm]) {
                    ExportAsTermTemplate(node, callback);
                };

            };

            // gel all templates
            obj.LoadTemplates = function () {

                var loaderId = "tb-data-toolbar-loader";
                handeler.loader("div#tb-templates-toolbar", "loader-tb-data", loaderId);
                spREST.Web.Lists.get(defs.L.Templates.Nm, defs.L.Templates.Q.getAllTemplates).done(function (data) {

                    var dataSet = [];
                    $.each(data.d.results, function (i, ele) {
                        dataSet.push({ Id: ele[defs.L.Templates.C.Id], tmp: ele[defs.L.Templates.C.Title], typ: ele[defs.L.Templates.C.TemplateType] });
                    });


                    loadDataToTemplateGrid(dataSet);

                }).fail(function (e) {
                    sputils.flashNotificationError(handeler.parseSPException(e));
                }).always(function () {
                    handeler.removeloader("#" + loaderId);
                });
            };

            // when click tempalte grid generate preview
            obj.clickTemplateGrid = function (rowIndex, rowData) {

                $("ul#templateTreePreview").empty();

                var loaderId = "tb-data-toolbar-loader";
                handeler.loader("div#tb-templates-toolbar", "loader-tb-data", loaderId);
                spREST.Web.Lists.get(defs.L.Templates.Nm, defs.L.Templates.Q.getTemplateById(rowData.Id)).done(function (data) {

                    var obj = JSON.parse(B64.decode(data.d.results[0][defs.L.Templates.C.TemplateData]));
                    loadExportMetaDataToTree("ul#templateTreePreview", obj.Data)

                }).fail(function (e) {
                    sputils.flashNotificationError(handeler.parseSPException(e));
                }).always(function () {
                    handeler.removeloader("#" + loaderId);
                });

            };

            function initTemplateApplyDialog() {
                $('div#dlg-apply-Templates').dialog('open');
            };

            function loadApplyTemplateDropDown(isRoot) {

                node = $('#templateTree').tree('getSelected');

                var loaderId = "tb-data-toolbar-loader";
                handeler.loader("div#tb-templates-toolbar", "loader-tb-data", loaderId);

                var temtype = "";

                if (isRoot) {
                    temtype = mapers.Template.Type.TermStore;
                }
                else {
                    if (node.attributes[mapers.TreeItem.Attributes.isTermStore]) {
                        temtype = mapers.Template.Type.Group;
                    }
                    else if (node.attributes[mapers.TreeItem.Attributes.isGroup]) {
                        temtype = mapers.Template.Type.TermSet;
                    }
                    else if (node.attributes[mapers.TreeItem.Attributes.isTermSet]) {
                        temtype = mapers.Template.Type.Term;
                    }
                    else if (node.attributes[mapers.TreeItem.Attributes.isTerm]) {
                        temtype = mapers.Template.Type.Term;
                    };
                }


                spREST.Web.Lists.get(defs.L.Templates.Nm, defs.L.Templates.Q.getTemplatesForType(temtype)).done(function (data) {

                    $("select#templatesToAplly").empty();

                    $.each(data.d.results, function (i, ele) {
                        $("select#templatesToAplly").append("<option value='" + ele[defs.L.Templates.C.Id] + "'>" + ele[defs.L.Templates.C.Title] + "</option>");
                    });

                    initTemplateApplyDialog();

                }).fail(function (e) {
                    sputils.flashNotificationError(handeler.parseSPException(e));
                }).always(function () {
                    handeler.removeloader("#" + loaderId);
                });
            }

            obj.loadDlgApply = function () {
                loadApplyTemplateDropDown(false);
            };

            obj.applyTemplate = function () {

                if ($("#from-apply-Templates").form('validate')) {

                    var id = $("select#templatesToAplly").val();

                    $('div#dlg-apply-Templates').dialog('close');
                    $.messager.progress({ msg: "Do not close the browser or refresh ", text: "Import in progress" });

                    spREST.Web.Lists.get(defs.L.Templates.Nm, defs.L.Templates.Q.getTemplateById(id)).done(function (data) {

                        var node = $('#templateTree').tree('getSelected');

                        var objr = new mapers.Template();
                        objr = JSON.parse(B64.decode(data.d.results[0][defs.L.Templates.C.TemplateData]));

                        var md = new spa.Metadata();
                        var tstoreId = getmmsid();

                        var obj;

                        switch (objr.Type) {
                            case mapers.Template.Type.Group:
                                var grp = md.getCreatedGroup(tstoreId, SP.Guid.newGuid().toString(), objr.Data.g[0].x);
                                grp = populateImportGroup(grp, objr.Data.g[0], objr.dl);

                                $.each(objr.Data.g[0].s, function (i, TS) {
                                    var ts = grp.createTermSet(TS.x, SP.Guid.newGuid().toString(), objr.dl);
                                    ts = populateImportTermSet(ts, objr.Data.g[0].s[i], objr.dl);

                                    // loop each term
                                    $.each(TS.t, function (j, T) {
                                        var t = ts.createTerm(T.x, objr.dl, SP.Guid.newGuid().toString());
                                        t = populateImportTerm(t, objr.Data.g[0].s[i].t[j], objr.dl);

                                        if (T.t.length > 0) {
                                            // Set options to overide id with new guid
                                            var op = {};
                                            op[enums.createChildTermsOptions.NewGuid] = true;
                                            createChildTerms(t, T, objr.dl, op);
                                        };
                                    });
                                });
                                obj = grp;
                                break;
                            case mapers.Template.Type.TermSet:
                                var termset = md.getCreatedTermSet(tstoreId, node.Id.toString(), SP.Guid.newGuid().toString(), objr.Data.g[0].s[0].x, objr.dl);
                                termset = populateImportTermSet(termset, objr.Data.g[0].s[0], objr.dl);

                                // loop each term
                                $.each(objr.Data.g[0].s[0].t, function (j, T) {
                                    var t = termset.createTerm(T.x, objr.dl, SP.Guid.newGuid().toString());
                                    t = populateImportTerm(t, objr.Data.g[0].s[0].t[j], objr.dl);

                                    if (T.t.length > 0) {
                                        // Set options to overide id with new guid
                                        var op = {};
                                        op[enums.createChildTermsOptions.NewGuid] = true;
                                        createChildTerms(t, objr.Data.g[0].s[0].t[j], objr.dl, op);
                                    };
                                });

                                obj = termset;

                                break;
                            case mapers.Template.Type.Term:

                                var term;
                                // check parent is term set
                                if (node.attributes[mapers.TreeItem.Attributes.isTermSet]) {
                                    term = md.getCreatedTerm(tstoreId, node.Id.toString(), SP.Guid.newGuid().toString(), objr.Data.g[0].s[0].t[0].x, objr.dl);
                                }
                                else { // parent is term, then child term chould be created
                                    var ptsid = node.attributes[mapers.TreeItem.Attributes.TermSetID].toString();
                                    term = md.getCreatedChildTerm(tstoreId, ptsid, node.Id.toString(), SP.Guid.newGuid().toString(), objr.Data.g[0].s[0].t[0].x, objr.dl);
                                };

                                term = populateImportTerm(term, objr.Data.g[0].s[0].t[0], objr.dl);

                                // loop each term
                                $.each(objr.Data.g[0].s[0].t[0].t, function (j, T) {

                                    var t = term.createTerm(T.x, objr.dl, SP.Guid.newGuid().toString());
                                    t = populateImportTerm(t, objr.Data.g[0].s[0].t[0].t[j], objr.dl);

                                    if (T.t.length > 0) {
                                        // Set options to overide id with new guid
                                        var op = {};
                                        op[enums.createChildTermsOptions.NewGuid] = true;
                                        createChildTerms(t, objr.Data.g[0].s[0].t[0].t[j], objr.dl, op);
                                    };
                                });


                                obj = term;

                                break;
                        };

                        spa.executeCtxAsync(obj).done(function () {
                            $.messager.alert('Import', defs.ImportOperationSuccess, 'info');
                            app.Default.loadMetadataTemplete(); // load the tree
                        }).fail(function (b) {
                            $.messager.alert(defs.String.Dlg.Title_Error, handeler.parseSPException(b), 'error');
                        }).always(function () {
                            $.messager.progress('close');
                        });
                    }).fail(function (e) {
                        $.messager.alert(defs.String.Dlg.Title_Error, handeler.parseSPException(b), 'error');
                    }).always(function () {
                        $.messager.progress('close');
                    });
                };
            };

            // load tab data
            obj.loadTab = function () {
                $("ul#templateTreePreview").empty(); // Clear the preview tree
                obj.LoadTemplates(); // Load the grid
                app.Default.loadMetadataTemplete(); // load the tree
            };

            return obj;

        })();

        return obj;
    })();

})(window.app = window.app || {}, jQuery);