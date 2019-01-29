/// <reference path="../scriptbase.js" />


(function (spa, $, undefined) {

    window._sa = spa;

    _sa.ctx = null;
    _sa.web = null;

    _sa.configs = {
        SPHostUrl: "",
        SPWebtUrl: ""
    };

    var configsettings = {
        paramSPHostUrl: "SPHostUrl",
        paramSPWebtUrl: "SPAppWebUrl"
    };

    init();

    utils.log(_sa.configs.SPHostUrl);
    utils.log(_sa.configs.SPWebtUrl);


    function saveConfigs() {

        var def = $.Deferred();

        var ctx = new SP.ClientContext.get_current();
        var web = ctx.get_web();
        ctx.load(web);
        var webProperties = web.get_allProperties();
        ctx.load(webProperties);
        webProperties.set_item("SPHostUrl", decodeURI(utils.getQueryStringParameter(configsettings.paramSPHostUrl)));
        webProperties.set_item("SPAppWebUrl", decodeURI(utils.getQueryStringParameter(configsettings.paramSPHostUrl)));
        web.update();
        _sa.configs.SPHostUrl = webProperties.get_item("SPHostUrl");
        _sa.configs.SPWebtUrl = webProperties.get_item("SPAppWebUrl");
        ctx.executeQueryAsync(
            function () { def.resolve(); },
            function (a, b) { def.reject(b.get_message); }
        );

        return def.promise();
    }

    function getConfigs() {

        var def = $.Deferred();

        var ctx = new SP.ClientContext.get_current();
        var web = ctx.get_web();
        ctx.load(web);
        var webProperties = web.get_allProperties();
        ctx.load(webProperties);
        ctx.executeQueryAsync(
            function () {
                try {
                    _sa.configs.SPHostUrl = webProperties.get_item("SPHostUrl");
                    _sa.configs.SPWebtUrl = webProperties.get_item("SPAppWebUrl");
                    def.resolve();
                } catch (e) {
                    def.reject(e.message);
                }
            },
            function (a, b) { def.reject(b.get_message); });

        return def.promise();
    }

    function init() {
        _sa.configs.SPHostUrl = decodeURIComponent(utils.getQueryStringParameter(configsettings.paramSPHostUrl));
        _sa.configs.SPWebtUrl = scriptbase.getAppWebUrlUrl();

        _sa.ctx = new SP.ClientContext.get_current();
        _sa.web = _sa.ctx.get_web();
    }

    // execute context or load and execute
    _sa.executeCtxAsync = function (object) {
        var def = $.Deferred();

        if (object) { _sa.ctx.load(object); };

        _sa.ctx.executeQueryAsync(function () {

            if (object) { def.resolve(object); }
            else { def.resolve(); };
        },
            function (a, b) {
                def.reject(b);
            });
        return def.promise();
    };

    _sa.Operations = function () {
    };

    _sa.Operations.prototype = {

        getListItemByIdAsync: function (listName, id) {

            var def = $.Deferred();
            var olist = _sa.web.get_lists().getByTitle(listName);
            var oitem = olist.getItemById(id);

            _sa.ctx.load(oitem);
            _sa.ctx.executeQueryAsync(function () {
                def.resolve(oitem);
            },
                function (a, b) {
                    def.reject(b);
                });
            return def.promise();
        },

        // Get by Caml (Name, Query, 'Include(Title,Type,ID,Modified)' )
        getListItemsByCAMLAsync: function (listName, query, fieldsFilter) {
            var def = $.Deferred();
            var ol = _sa.web.get_lists().getByTitle(listName);
            var qry = new SP.CamlQuery();
            qry.set_viewXml(query);
            var items = ol.getItems(qry);

            _sa.ctx.load(items, fieldsFilter);
            _sa.ctx.executeQueryAsync(function () {
                def.resolve(items);
            },
                function (a, b) {
                    def.reject(b);
                });
            return def.promise();
        },

        // Get reference to new list Item 
        getNewListItem: function (listName) {

            var olist = _sa.web.get_lists().getByTitle(listName);
            var itemCreateInfo = new SP.ListItemCreationInformation();
            return olist.addItem(itemCreateInfo);
        },

        // Get reference to the List Item
        getListItem: function (listName, Id) {
            var olist = _sa.web.get_lists().getByTitle(listName);
            return olist.getItemById(Id);
        },

        //insert and Update the item
        insertListItem: function (listItem) {
            var def = $.Deferred();
            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(listItem);
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },

        updateListColumnByIdAsync: function (listName, id, columnsName, value) {

            var def = $.Deferred();
            var olist = _sa.web.get_lists().getByTitle(listName);
            var oitem = olist.getItemById(id);
            oitem.set_item(columnsName, value);
            oitem.update();
            _sa.ctx.executeQueryAsync(function () {
                def.resolve(oitem);
            },
                function (a, b) {
                    def.reject(b);
                });
            return def.promise();
        },

        deleteListItemByIdAsync: function (listName, id) {

            var def = $.Deferred();
            var olist = _sa.web.get_lists().getByTitle(listName);
            olist.getItemById(id).deleteObject();
            _sa.ctx.executeQueryAsync(function () {
                def.resolve();
            },
                function (a, b) {
                    def.reject(b);
                });
            return def.promise();
        },

        // Create a file in Document Library
        createFileinDocumentLibraryAsync: function (listName, fileName, ContentAsString) {
            var def = $.Deferred();
            var ol = _sa.web.get_lists().getByTitle(listName);

            var fci = new SP.FileCreationInformation();
            fci.set_url(fileName);

            var content = new SP.Base64EncodedByteArray();
            for (var i = 0; i < ContentAsString.length; i++) {
                content.append(ContentAsString.charCodeAt(i));
            }

            fci.set_content(content);

            var f = ol.get_rootFolder().get_files().add(fci);

            _sa.ctx.load(f);
            _sa.ctx.executeQueryAsync(function () {
                def.resolve(f);
            },
                function (a, b) {
                    def.reject(b);
                });
            return def.promise();
        },
        // Read a file as a string
        getFileDatainDocumentLibraryAsync: function (listname, filename) {

            var def = $.Deferred();

            var clientContext = new SP.ClientContext.get_current();
            var oWebsite = clientContext.get_web();

            clientContext.load(oWebsite);
            clientContext.executeQueryAsync(function () {
                fileUrl = oWebsite.get_serverRelativeUrl() + "/Lists/" + listname + "/" + filename;
                $.ajax({ url: fileUrl, type: "GET" })
                    .done(function (data) {
                        def.resolve(data);
                    })
                    .error(function () {
                        def.reject(false)
                    });
            }), function (a, b) {
                def.reject(b);
            };

            return def.promise();
        }

    };

    _sa.Permisions = function () {
    };

    _sa.Permisions.prototype = {
        doesUserHaveManagewebAsync: function () {

            var def = $.Deferred();

            var permis = new SP.BasePermissions();
            permis.set(SP.PermissionKind.manageWeb);

            var res = _sa.web.doesUserHavePermissions(permis);

            _sa.ctx.executeQueryAsync(function () {
                if (res.get_value()) {
                    def.resolve();
                } else {
                    def.reject();
                }

            },
                function (a, b) {
                    def.reject(b);
                });
            return def.promise();
        }
    };

    //use  $$.IncludeScript($$.js.spTaxonomy);
    _sa.Metadata = function () {

    };

    _sa.Metadata.prototype = {
        // get all term stores
        getTermStoresAsync: function () {
            var def = $.Deferred();

            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termStores = taxSession.get_termStores();

            _sa.ctx.load(termStores);
            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(termStores);
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },
        // get all term stores
        getTermStoreByIdAsync: function (id) {
            var def = $.Deferred();

            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termStore = taxSession.get_termStores().getById(id);

            _sa.ctx.load(termStore);
            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(termStore);
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },
        //get all groups for a term store
        getTermGroupsbyStoreIdAsync: function (Id) {
            var def = $.Deferred();

            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termStoresGroups = taxSession.get_termStores().getById(Id).get_groups();

            _sa.ctx.load(termStoresGroups);
            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(termStoresGroups);
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },
        // get group under term store
        getTermGroupbyIdAsync: function (storeId, groupId) {
            var def = $.Deferred();

            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termGroup = taxSession.get_termStores().getById(storeId).get_groups().getById(groupId);

            _sa.ctx.load(termGroup);
            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(termGroup);
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },
        //Delete Group
        deleteTermGroupbyIdAsync: function (storeId, groupId) {
            var def = $.Deferred();

            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termGroup = taxSession.get_termStores().getById(storeId).get_groups().getById(groupId);
            termGroup.deleteObject();
            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(true);
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },
        // creare a group under termstore
        createTermGroupAsync: function (storeId, groupName, groupId) {
            var def = $.Deferred();

            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termStore = taxSession.get_termStores().getById(storeId);
            var termGroup = termStore.createGroup(groupName, groupId);

            _sa.ctx.load(termGroup);
            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve();
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },
        //get termset by id
        getTermSetbyIdAsync: function (storeId, groupId, termsetId) {

            var def = $.Deferred();

            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termGroup = taxSession.get_termStores().getById(storeId).get_groups().getById(groupId);
            var termSet = termGroup.get_termSets().getById(termsetId);

            _sa.ctx.load(termSet);
            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(termSet);
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();

        },
        // get all termsetes for group
        getallTermSetsbyGroupAsync: function (storeId, groupId) {
            var def = $.Deferred();

            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termGroup = taxSession.get_termStores().getById(storeId).get_groups().getById(groupId);
            var termsets = termGroup.get_termSets();
            _sa.ctx.load(termsets);
            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(termsets);
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },
        //Delete term set
        deleteTermSetAsync: function (storeId, termSetId) {
            var def = $.Deferred();

            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termStore = taxSession.get_termStores().getById(storeId);
            var termSet = termStore.getTermSet(termSetId);
            termSet.deleteObject();
            _sa.ctx.executeQueryAsync(
               function () {
                   def.resolve();
               },
               function (a, b) {
                   def.reject(b);
               }
           );
            return def.promise();
        },
        //get referenc to child term
        getCreatedChildTerm: function (storeId, termSetId, parenttermId, termId, termName, lcid) {
            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var parentTerm = taxSession.get_termStores().getById(storeId).getTermSet(termSetId).getTerm(parenttermId);
            return parentTerm.createTerm(termName, lcid, termId);
        },
        //get reference to term
        getCreatedTerm: function (storeId, termSetId, termId, termName, lcid) {
            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            return taxSession.get_termStores().getById(storeId).getTermSet(termSetId).createTerm(termName, lcid, termId);
        },
        //get a reference to termSet
        getCreatedTermSet: function (storeId, groupId, termsetId, termsetName, lcid) {
            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termGroup = taxSession.get_termStores().getById(storeId).get_groups().getById(groupId);
            return termGroup.createTermSet(termsetName, termsetId, lcid)
        },
        // Get reference to created group
        getCreatedGroup: function (storeId, groupId, groupName) {
            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            return taxSession.get_termStores().getById(storeId).createGroup(groupName, groupId);
        },
        // create term set in with chids
        createTermSetWithChildsAsync: function (termSet) {
            var def = $.Deferred();
            _sa.ctx.load(termSet);
            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(termSet);
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },
        // create term set in a group
        createTermSetAsync: function (storeId, groupId, termsetId, termsetName, lcid) {
            var def = $.Deferred();

            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termGroup = taxSession.get_termStores().getById(storeId).get_groups().getById(groupId);
            var termSet = termGroup.createTermSet(termsetName, termsetId, lcid)
            _sa.ctx.load(termSet);
            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(termSet);
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },
        // Create Group with childs
        createGroupWithChildsAsync: function (group) {
            var def = $.Deferred();
            _sa.ctx.load(group);
            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(group);
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },
        //get all terms in a term set
        getallTermsAsync: function (storeId, groupId, termsetId) {
            var def = $.Deferred();

            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termGroup = taxSession.get_termStores().getById(storeId).get_groups().getById(groupId);
            var terms = termGroup.get_termSets().getById(termsetId).getAllTerms();
            _sa.ctx.load(terms);
            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(terms);
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },
        //get terms in a term set
        getTermsAsync: function (storeId, groupId, termsetId) {
            var def = $.Deferred();

            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termGroup = taxSession.get_termStores().getById(storeId).get_groups().getById(groupId);
            var terms = termGroup.get_termSets().getById(termsetId).get_terms();
            _sa.ctx.load(terms);
            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(terms);
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },
        //get term by id
        getTermByIdAsync: function (storeId, groupId, termsetId, termId) {
            var def = $.Deferred();
            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termGroup = taxSession.get_termStores().getById(storeId).get_groups().getById(groupId);
            var term = termGroup.get_termSets().getById(termsetId).getTerm(termId);

            _sa.ctx.load(term);
            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(term);
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },
        //get term by id with labels
        getTermByIdAsyncWithLabels: function (storeId, groupId, termsetId, termId, lcid) {
            var def = $.Deferred();
            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termGroup = taxSession.get_termStores().getById(storeId).get_groups().getById(groupId);
            var term = termGroup.get_termSets().getById(termsetId).getTerm(termId);
            var alllabels = term.getAllLabels(lcid)

            _sa.ctx.load(alllabels);
            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(alllabels);
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },
        //get all childs for a term
        getChildTermsAsync: function (storeId, groupId, termsetId, termId) {
            var def = $.Deferred();

            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termGroup = taxSession.get_termStores().getById(storeId).get_groups().getById(groupId);
            var terms = termGroup.get_termSets().getById(termsetId).getTerm(termId).get_terms();

            _sa.ctx.load(terms);
            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(terms);
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },
        // set term description by Id
        setTermDescriptionById: function (storeId, groupId, termsetId, termId, lcid, desc) {
            var def = $.Deferred();
            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termGroup = taxSession.get_termStores().getById(storeId).get_groups().getById(groupId);
            var term = termGroup.get_termSets().getById(termsetId).getTerm(termId);
            term.setDescription(desc, lcid);

            _sa.ctx.load(term);
            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(term);
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },
        // set term Name by Id
        setTermNameById: function (storeId, groupId, termsetId, termId, lcid, name) {
            var def = $.Deferred();
            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termGroup = taxSession.get_termStores().getById(storeId).get_groups().getById(groupId);
            var term = termGroup.get_termSets().getById(termsetId).getTerm(termId);
            term.set_name(name, lcid);

            _sa.ctx.load(term);
            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(term);
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },
        // set term Tagging by Id
        setTermTaggingById: function (storeId, groupId, termsetId, termId, tagging) {
            var def = $.Deferred();
            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termGroup = taxSession.get_termStores().getById(storeId).get_groups().getById(groupId);
            var term = termGroup.get_termSets().getById(termsetId).getTerm(termId);
            term.set_isAvailableForTagging(tagging);

            _sa.ctx.load(term);
            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(term);
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },
        // create a term under termSet
        createTermAsync: function (storeId, termsetId, termId, termName, lcid) {
            var def = $.Deferred();

            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termSet = taxSession.get_termStores().getById(storeId).getTermSet(termsetId);
            var term = termSet.createTerm(termName, lcid, termId);
            _sa.ctx.load(term);
            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(term);
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },
        //Delete Term
        deleteTermByIdAsync: function (storeId, groupId, termsetId, termId) {
            var def = $.Deferred();
            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termGroup = taxSession.get_termStores().getById(storeId).get_groups().getById(groupId);
            var term = termGroup.get_termSets().getById(termsetId).getTerm(termId);
            term.deleteObject();

            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(true);
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },
        //Create child term
        createChildTermAsync: function (storeId, termsetId, parenttermId, termId, termName, lcid) {
            var def = $.Deferred();

            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var parentTerm = taxSession.get_termStores().getById(storeId).getTermSet(termsetId).getTerm(parenttermId);
            var term = parentTerm.createTerm(termName, lcid, termId);

            _sa.ctx.load(term);
            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(term);
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },

        // Create Term Lable
        createTermLableAsync: function (storeId, groupId, termsetId, termId, lcid, lable, isDefault) {
            var def = $.Deferred();
            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termGroup = taxSession.get_termStores().getById(storeId).get_groups().getById(groupId);
            var term = termGroup.get_termSets().getById(termsetId).getTerm(termId);

            if (isDefault) {
                isDefault = false;
            }

            term.createLabel(lable, lcid, isDefault)

            _sa.ctx.load(term);
            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(term);
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },
        // Delete Lable
        deletelableByAsync: function (storeId, groupId, termsetId, termId, lcid, lable) {
            var def = $.Deferred();
            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termGroup = taxSession.get_termStores().getById(storeId).get_groups().getById(groupId);
            var lbl = termGroup.get_termSets().getById(termsetId).getTerm(termId).getAllLabels(lcid).getByValue(lable);
            lbl.deleteObject();

            _sa.ctx.load(lbl);
            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(true);
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },
        // Term term Shared properties
        getTermSharedPropertiesByAsync: function (storeId, groupId, termsetId, termId) {

            var def = $.Deferred();
            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termGroup = taxSession.get_termStores().getById(storeId).get_groups().getById(groupId);
            var tm = termGroup.get_termSets().getById(termsetId).getTerm(termId);

            _sa.ctx.load(tm, 'CustomProperties');
            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(tm.get_customProperties());
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },
        // Create Custom Shared Property
        createSharedCustomPropertyByAsync: function (storeId, groupId, termsetId, termId, name, value) {
            var def = $.Deferred();
            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termGroup = taxSession.get_termStores().getById(storeId).get_groups().getById(groupId);
            termGroup.get_termSets().getById(termsetId).getTerm(termId).setCustomProperty(name, value);

            //_sa.ctx.load(cp);
            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(true);
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },
        // delete shared custom property
        deleteSharedCustomPropertyByAsync: function (storeId, groupId, termsetId, termId, name) {
            var def = $.Deferred();
            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termGroup = taxSession.get_termStores().getById(storeId).get_groups().getById(groupId);
            termGroup.get_termSets().getById(termsetId).getTerm(termId).deleteCustomProperty(name);

            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(true);
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },
        // Get term Local properties
        getTermLocalPropertiesByAsync: function (storeId, groupId, termsetId, termId) {
            var def = $.Deferred();
            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termGroup = taxSession.get_termStores().getById(storeId).get_groups().getById(groupId);
            var tm = termGroup.get_termSets().getById(termsetId).getTerm(termId);

            _sa.ctx.load(tm, 'LocalCustomProperties');
            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(tm.get_localCustomProperties());
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },
        // Get Local and Shared Term Properties
        getTermLocalSharedPropertiesByTermIdAsync: function (storeId, groupId, termsetId, termId) {
            var def = $.Deferred();
            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termGroup = taxSession.get_termStores().getById(storeId).get_groups().getById(groupId);
            var tm = termGroup.get_termSets().getById(termsetId).getTerm(termId);

            _sa.ctx.load(tm, 'LocalCustomProperties', 'CustomProperties');
            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(tm.get_localCustomProperties(), tm.get_customProperties());
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },
        // Create Custom Local Property
        createLocalCustomPropertyByAsync: function (storeId, groupId, termsetId, termId, name, value) {
            var def = $.Deferred();
            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termGroup = taxSession.get_termStores().getById(storeId).get_groups().getById(groupId);
            termGroup.get_termSets().getById(termsetId).getTerm(termId).setLocalCustomProperty(name, value);

            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(true);
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        },
        // delete local custom property
        deleteLocalCustomPropertyByAsync: function (storeId, groupId, termsetId, termId, name) {
            var def = $.Deferred();
            var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(_sa.ctx);
            var termGroup = taxSession.get_termStores().getById(storeId).get_groups().getById(groupId);
            termGroup.get_termSets().getById(termsetId).getTerm(termId).deleteLocalCustomProperty(name);

            _sa.ctx.executeQueryAsync(
                function () {
                    def.resolve(true);
                },
                function (a, b) {
                    def.reject(b);
                }
            );
            return def.promise();
        }

    };

}(window.spa = window.spa || {}, jQuery));

(function (sputils, $, undefined) {

    sputils.flashNotificationWarninig = function (message) {
        var status = SP.UI.Status.addStatus("Warning :", message);
        SP.UI.Status.setStatusPriColor(status, 'yellow');
        setTimeout(function () { SP.UI.Status.removeStatus(status); }, 3000);
    };

    sputils.flashNotificationInfo = function (message) {
        var status = SP.UI.Status.addStatus("Info: ", message);
        SP.UI.Status.setStatusPriColor(status, 'blue');
        setTimeout(function () { SP.UI.Status.removeStatus(status); }, 3000);
    };

    sputils.flashNotificationError = function (message) {
        var status = SP.UI.Status.addStatus("Error :", message);
        SP.UI.Status.setStatusPriColor(status, 'red');
        setTimeout(function () { SP.UI.Status.removeStatus(status); }, 3000);
    }

    sputils.SPModalDialog = function (dialogUrl, dialogTitle, callback, width, height) {
        var options =
        {
            url: $$.getAppWebUrlUrl() + dialogUrl,
            autoSize: true,
            allowMaximize: true,
            showClose: true,
            title: dialogTitle,
            dialogReturnValueCallback: function (dialogResult) {
                if (callback) {
                    callback(dialogResult);
                }
            }
        };

        if (width) {
            options.width = width;
        }

        if (height) {
            options.height = height;
        }

        SP.UI.ModalDialog.showModalDialog(options);
    };

    sputils.initializePeoplePicker = function (peoplePickerElementId) {

        var schema = {};
        schema['PrincipalAccountType'] = 'User,DL,SecGroup,SPGroup';
        schema['SearchPrincipalSource'] = 15;
        schema['ResolvePrincipalSource'] = 15;
        schema['AllowMultipleValues'] = true;
        schema['MaximumEntitySuggestions'] = 50;
        schema['Width'] = '280px';

        SPClientPeoplePicker_InitStandaloneControlWrapper(peoplePickerElementId, null, schema);
    };

    sputils.setPeoplePickerValues = function (peoplePickerElementId, itemsArray) {

        if (itemsArray) {
            //peoplePickerHosting_TopSpan
            var spPeoplePicker = SPClientPeoplePicker.SPClientPeoplePickerDict[peoplePickerElementId + "_TopSpan"];
            var spPeoplePickerEditorID = "#" + peoplePickerElementId + "_TopSpan_EditorInput";

            $.each(itemsArray, function (i, ele) {
                $(spPeoplePickerEditorID).val(ele.get_lookupValue());
                spPeoplePicker.AddUnresolvedUserFromEditor(true);
            });
        }


    };

    sputils.PeoplePickerClear = function (peoplePickerElementId) {

        var spPeoplePickerEditorID = "#" + peoplePickerElementId + "_TopSpan_ResolvedList";
        $(spPeoplePickerEditorID).empty();
    };

    sputils.setPeoplePickerEnable = function (peoplePickerElementId, enable) {

        //peoplePickerHosting_TopSpan
        var spPeoplePicker = SPClientPeoplePicker.SPClientPeoplePickerDict[peoplePickerElementId + "_TopSpan"];
        spPeoplePicker.SetEnabledState(enable);
    };

}(window.sputils = window.sputils || {}, jQuery));

(function (spREST, $, undefined) {

    function getListItemType(listname) {
        return "SP.Data." + listname[0].toUpperCase() + listname.substring(1) + "ListItem";
    }

    // Web REST
    spREST.Web = spREST.Web || {};

    // Web RegionalSettings
    spREST.Web.RegionalSettings = spREST.Web.RegionalSettings || {};

    spREST.Web.RegionalSettings.TimeZone = function () {
        var def = $.Deferred();
        $.ajax({
            url: $$.getAppWebUrlUrl() + "/_api/web/RegionalSettings/timeZone",
            dataType: 'xml',
            success: function (data) {
                def.resolve(data);
            },
            error: function (data) {
                def.reject(data.statusText);
            }
        });

        return def.promise();
    };

    // Web/List REST
    spREST.Web.Lists = spREST.Web.Lists || {};

    spREST.Web.Lists.add = function (listname, data) {

        var def = $.Deferred();

        var item = $.extend({
            "__metadata": { "type": getListItemType(listname) }
        }, data);

        $.ajax({
            url: $$.getAppWebUrlUrl() + "/_api/web/lists/getbytitle('" + listname + "')/items",
            type: "POST",
            contentType: "application/json;odata=verbose",
            data: JSON.stringify(item),
            headers: {
                "Accept": "application/json;odata=verbose",
                "X-RequestDigest": $("#__REQUESTDIGEST").val()
            },
            success: function (data) {
                def.resolve(data);
            },
            error: function (data) {
                def.reject(data.statusText);
            }
        });


        return def.promise();

    };

    spREST.Web.Lists.get = function (listname, Query) {

        var def = $.Deferred();

        var u = $$.getAppWebUrlUrl() + "/_vti_bin/listdata.svc/" + listname;

        if (Query) {
            u += Query;
        }

        $.ajax({
            dataType: "json",
            url: u,
            success: function (data) {
                def.resolve(data);
            },
            error: function (data) {
                def.reject(data.statusText);
            }
        });

        return def.promise();

    };


}(window.spREST = window.spREST || {}, jQuery));