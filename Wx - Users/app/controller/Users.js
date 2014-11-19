Ext.define('wx.controller.Users', {
    require: [ 'Ext.form.FormPanel', 'Ext.Window' ],
    extend: 'Ext.app.Controller',
    views:  [ 'common.UserEdit', 'common.UserList' ],
    stores: [ 'Users' ],
    models: [ 'User'],

    refs: [
        { ref: 'useredit', selector: 'useredit'},
        { ref: 'editForm', selector: 'useredit form'},
        { ref: 'userlist', selector: 'userlist' },
        { ref: 'userstore', selector: 'userstore' },
        { ref: 'wxworkspacecenter', selector: 'wxworkspacecenter' },
        { ref: 'wxworkspacefooter', selector: 'wxworkspacefooter' }
    ],

    constructor: function(config) {

        //var msg = 'Starting ...';
        //WxLog("debug", "controller/Users.constructor", msg);

        try {

            this.callParent([config]);

            this.initialized = false;
            this.displaying = false;

            this.initialLoadCalled = false;

        }

        catch(ex) {
            WxLog("error", "controller/Users.constructor", ex.message);
        }
        
    },

    init: function() {

        //var msg = 'Starting ...';
        //WxLog("debug", "controller/Users.init", msg);

        try {

            if( !this.initialized ){

                this.control({

                    "#usertabpanel":{
                      tabchange: function(pan, nc, oc, opts){
                          this.tabChangeProcess(pan, nc, oc, opts);
                      }
                    },
                    "useredit button[action=formApply]": {
                        click: function() {
                            this.userApply(false);
                        }
                    },
                    "useredit button[action=formCopy]": {
                        click: function() {
                            this.userCopy();
                        }
                    },
                    "useredit button[action=formCancel]": {
                        click: function() {
                            this.userCancel();
                        }
                    },
                    "useredit button[action=formDelete]": {
                        click: function() {
                            this.userDelete();
                        }
                    },
                    "useredit button[action=formNew]": {
                        click: function() {
                            this.userNew();
                        }
                    },
                    "useredit button[action=formOk]": {
                        click: function() {
                            this.userApply(true);
                        }
                    }
                });

                document.app.addListener( 'logoutEvent', this.logoutEvent, this );

                this.initialized = true;

                if ( this.displaying ) {
                    this.show();
                }

            }

        }
        catch(ex) {
            WxLog("error", "controller/Users.init", ex.message);
        }
    }, //init
    
    show: function() {

        //var msg = 'Starting ...';
        //WxLog("debug", "controller/Users.show", msg);

        try {

            this.resumeComponentEvents();
                
            if ( !this.initialized ) {

                //msg = 'The form is not initialized';
                //WxLog("debug", "controller/Users.show", msg);

                this.displaying = true;
                this.init();

            } else {

                //msg = 'The form is already initialized';
                //WxLog("debug", "controller/Users.show", msg);


                //  ---  Edit Form Area ---

                document.app.userstate = 'showing';
                
                var editComponent = Ext.getCmp('UserEditID');
                var editWorkspace = Ext.getCmp('wxworkspacecenter');
                
                if(isDefined(editComponent) ){
                    editComponent.show();
                } else {
                    editComponent = Ext.create('wx.view.common.UserEdit');
                }
                editWorkspace.setHeight(editComponent.height);
                editWorkspace.up().doLayout();


                //  ---  List Area ---

                var listComponent = Ext.getCmp('UserListID');
                var listWorkspace = Ext.getCmp('wxworkspacefooter');

                if(isDefined(listComponent)){
                    listComponent.show();
                } else {
                    listComponent = Ext.create('wx.view.common.UserList');
                }
                listWorkspace.setHeight(listComponent.height);
                listWorkspace.up().doLayout();

                //msg = 'Finished setting up components for show';
                //WxLog("debug", "controller/Audits.show", msg);


                // --- Reload the store(s) as appropriate ---

                var editForm = Ext.getCmp('UserEditFormID'); //undefined when initially loading the data

                if(isDefined(editForm)){

                    if( !this.initialLoadCalled ) {
                        this.loadSpecialCombos();
                    } else {
                        this.initialLoadCall(true);
                    }

                }
                
                this.displaying = false;
                
            }
        }
        catch(ex) {
            WxLog("error", "controller/Users.show", ex.message);
        }
    }, //show


    loadSpecialCombos: function(){
        //Used to load all the enumcombobox stores on INITIAL form render

        //var msg = 'Starting ...';
        //WxLog("debug", "controller/Users.loadSpecialCombos()", msg);

        try {

            var theCombosList = [ 'UserEditAgencyCombo',
                                  'UserEditUserPrefixCombo',
                                  'UserEditUserSuffixCombo',
                                  'UserEditStateCombo',
                                  'UserEditPrivilegeLevelCombo' ];

            var i=0;
            for(i=0;i<theCombosList.length;i++){
                Ext.getCmp(theCombosList[i]).store.load();
            }

            //msg = 'Finished';
            //WxLog("debug", "controller/Users.loadSpecialCombos()", msg);

        }
        catch(ex) {
            WxLog("error", "controller/Users.loadSpecialCombos()", ex.message);
            Ext.getBody().unmask();
        }

    }, //loadSpecialCombos

    initialLoadCall: function( forceStoreLoad ){

        try {

            forceStoreLoad = ( isDefined(forceStoreLoad) && ( forceStoreLoad == true || forceStoreLoad == false ) ) ? forceStoreLoad : false;

            if( this.initialLoadCalled && !forceStoreLoad ){
                return; //Added to increase speed
            }

            var theCombosList = [ 'UserEditAgencyCombo',
                                  'UserEditUserPrefixCombo',
                                  'UserEditUserSuffixCombo',
                                  'UserEditStateCombo',
                                  'UserEditPrivilegeLevelCombo'  ];

            var i=0;
            var loaded=true;

            for(i=0;i<theCombosList.length;i++){

                if(Ext.getCmp(theCombosList[i]).store.isLoading()){
                    loaded = false;
                    break;
                }

            }

            if( ( loaded && !this.initialLoadCalled) || forceStoreLoad ){

                this.initialLoadCalled = true;

                Ext.getStore('Users').load();

            }

        }
        catch(ex) {
            WxLog("error", "controller/Users.initialLoadCall", ex.message);
        }

    }, //initialLoadCall




    hide: function() {

        //var msg = 'Starting ...';
        //WxLog("debug", "controller/Users.hide", msg);

        try {

            this.clear(true);

            document.app.userstate = 'hiding';

            var editComponent = Ext.getCmp('UserEditID');
            editComponent.hide();

            var editWorkspace = Ext.getCmp('wxworkspacecenter');
            editWorkspace.up().doLayout();

            var listComponent = Ext.getCmp('UserListID');
            listComponent.hide();

            var listWorkspace = Ext.getCmp('wxworkspacefooter');
            listWorkspace.up().doLayout();

            var auditDock = document.app.getController('AuditDock');
            auditDock.hidingSelection('Users');

            //msg = 'Finished';
            //WxLog("debug", "controller/Users.hide", msg);

        }
        catch(ex) {
            WxLog("error", "controller/Users.hide", ex.message);
        }
    }, //hide

    clear: function(clearList) {

        //var msg = "Starting ...";
        //WxLog("debug", "controller/Users.clear", msg);

        try {

            clearList = ( isDefined(clearList) && ( clearList == true || clearList == false ) ) ? clearList : false;

            var theEditForm = Ext.getCmp('UserEditFormID').getForm();
            var theFormModel = Ext.create('wx.model.User');

            theEditForm.loadRecord(theFormModel);

            Ext.getCmp('user__user_name_image').setSrc('');

            var theCboStore = Ext.getCmp('UserEditAgencyCombo').getStore();
            theCboStore.clearFilter(true);

            theEditForm.findField('user__new_password').setValue(null);
            Ext.getCmp('user__new_password_image').setSrc('');

            theEditForm.findField('user__new_password_confirm').setValue(null);
            Ext.getCmp('user__new_password_confirm_image').setSrc('');

            theEditForm.findField('user__is_copy').setValue(false);
            theEditForm.findField('user__is_new').setValue(false);

            this.enableAddEdit();

            theEditForm.clearInvalid();


            if ( clearList == true ) {
                //console.log('Clearing the list');

                var theListGrid = Ext.getCmp('UserListID');
                theListGrid.getStore().load(theFormModel);

            }

            //msg = 'Finished';
            //WxLog("debug", "controller/Users.clear", msg);

        }
        catch(ex) {
            WxLog("error", "controller/Users.clear", ex.message);
        }
    }, //clear

    logoutEvent: function(agency_id, user_id) {

        var msg = ' logout : ' + agency_id.toString() + ', type: ' + user_id.toString();
        WxLog("debug", "controller/Users.logoutEvent", msg);

        try {

            this.hide();

        }
        catch(ex)
        {
            WxLog("error", "controller/Users.logoutEvent", ex.message);
        }
    }, //logoutEvent

    userCancel: function () {

        //var msg = 'Starting ...';
        //WxLog("debug", "controller/Users.userCancel", msg);

        try {

            var theEditForm = Ext.getCmp('UserEditFormID').getForm();

            var isCopy = stringToBoolean(theEditForm.findField('user__is_copy').getValue());
            var isNew = stringToBoolean(theEditForm.findField('user__is_new').getValue());

            if( isCopy || isNew ){

                var theStore = Ext.getStore('Users');

                if( theStore.getCount() > 0 ){
                    var theListObj = Ext.getCmp('UserListID');
                    theListObj.getSelectionModel().select(0, false, false);
                } else {
                    this.clear();
                }

            } else {
                this.hide();
            }

        }
        catch(ex) {
            WxLog("error", "controller/Users.userCancel", ex.message);
        }
    }, //userCancel




    onUsersStoreLoad: function( theStore ) {

        //var msg = 'Starting ...';
        //WxLog("debug", "controller/Users.onUsersStoreLoad", msg);

        if( theStore.getCount() > 0 ){

            var theEditForm = Ext.getCmp('UserEditFormID').getForm();
            var theListGrid = Ext.getCmp('UserListID');
            var theGridStore = theListGrid.getStore();

            var theListIdx = 0;
            var tmpListIdx = 0;

            var theUserID = parseInt(theEditForm.findField('user__selected_user_id').getValue());
            theUserID = ( isNaN(theUserID) || theUserID == -1 ) ? 0 : theUserID;

            if( theUserID == 0 ) {

                var theUserName = theEditForm.findField('user__selected_user_name').getValue();

                if ( isDefined(theUserName) && theUserName !=='' ) {

                    tmpListIdx = theGridStore.findExact( 'user__user_name', theUserName );

                    if( tmpListIdx > -1 ){ theListIdx = tmpListIdx; }

                }

            } else {

                tmpListIdx = theGridStore.findExact( 'user__user_id', theUserID );

                if( tmpListIdx > -1 ){ theListIdx = tmpListIdx; }

            }

            theListGrid.getSelectionModel().select(theListIdx, false, false);

        }


    }, //onUsersStoreLoad

    onUserListChange: function( theRecord ) {

        //var msg = 'Starting ...';
        //WxLog("debug", "controller/Users.onUserListChange", msg);

        try {

            // Basic form config ...

            var theEditFormParent = Ext.getCmp('UserEditFormID');
            var theEditForm = theEditFormParent.getForm();

            Ext.getCmp('userLoginInfo').setDisabled(false);
            Ext.getCmp('userPersonInfo').setDisabled(false);
            Ext.getCmp('userContactInfo').setDisabled(false);

            theEditForm.findField('user__is_copy').setValue(false);
            theEditForm.findField('user__is_new').setValue(false);

            Ext.getCmp('user__user_name_image').setSrc('');

            theEditForm.findField('user__new_password').setValue(null);
            Ext.getCmp('user__new_password_image').setSrc('');

            theEditForm.findField('user__new_password_confirm').setValue(null);
            Ext.getCmp('user__new_password_confirm_image').setSrc('');

            var thePrivCboStore = Ext.getCmp('UserEditPrivilegeLevelCombo').getStore();
            thePrivCboStore.clearFilter(true);

            var theAgencyCboStore = Ext.getCmp('UserEditAgencyCombo').getStore();
            theAgencyCboStore.clearFilter(true);


            var theListGrid = Ext.getCmp('UserListID');

            // Load the selected user
            if( theListGrid.getSelectionModel().hasSelection() ){
                var selObj = theListGrid.getSelectionModel().getSelection()[0];
                theEditFormParent.loadRecord(selObj);
            }


            this.enableAddEdit();
            this.filterPrivlegeLevelCombo();
            this.filterAgencyCombo();

        }
        catch(ex) {
            WxLog("error", "controller/Users.onUserListChange", ex.message);
        }

    }, //onUserListChange



    enableAddEdit: function () {

        try{

            var newBtn = Ext.getCmp('userNewButton');
            var copyBtn = Ext.getCmp('userCopyButton');
            var deleteBtn = Ext.getCmp('userDeleteButton');
            var applyBtn = Ext.getCmp('userApplyButton');
            var okBtn = Ext.getCmp('userOkButton');

            var canCreate = false;
            var canCopy = false;
            var canDelete = false;
            var canEdit = false;

            var isSelf = false;
            var authIsAdmin = false;

            if( isDefined(document.app.user_id) && isDefined(document.app.userAuths) ){


                isSelf = ( Ext.getCmp('user__user_id').getValue() == document.app.user_id );


                var privLevelCode = document.app.userAuths['userInfo'].privLevelCode;

                if( privLevelCode == 'site_admin' ){

                    authIsAdmin = true;

                } else if( privLevelCode == 'state_admin' ){

                    var authStateId = parseInt(document.app.userAuths['userInfo'].state_id, 10);
                    var userStateId = parseInt(Ext.getCmp('UserEditFormID').getForm().findField('user__state_id').getValue(), 10);

                    authIsAdmin = ( authStateId == userStateId );

                } else if( privLevelCode == 'agency_admin' ){

                    var userAgencyId = parseInt(Ext.getCmp('UserEditFormID').getForm().findField('user__agency_id').getValue(), 10);

                    //var authAgencyId = parseInt(document.app.userAuths['userInfo'].state_id, 10);

                    var theAuthUsersAgencies = [];
                    var aCtr = 0;
                    for ( aCtr = 0; aCtr < document.app.userAuths['userAgencies'].length; aCtr++ ) {
                        theAuthUsersAgencies.push( parseInt(document.app.userAuths['userAgencies'][aCtr].id) );
                    }

                    authIsAdmin = ( theAuthUsersAgencies.indexOf(userAgencyId) >= 0 ); // agencyIsInAuthList == true

                }


                canCreate = userAuthCheck( 'user', 'create', false );

                //canDelete = userAuthCheck( 'user', 'edit', isSelf );
                canDelete = canCreate;

                canCreate = userAuthCheck( 'user', 'create', isSelf );
                var userPrivLevelId = parseInt(Ext.getCmp('UserEditPrivilegeLevelCombo').getValue());
                var isGlobalAdmin = ( userPrivLevelId <= 2 );
                canCopy = canCreate && !isGlobalAdmin;

                canEdit = userAuthCheck( 'user', 'edit', isSelf );


            }

            //console.log( 'controller/Users.enableAddEdit: isSelf = ', isSelf, ', authIsAdmin = ', authIsAdmin, ', canCreate = ', canCreate, ', canCopy = ', canCopy, ', canDelete = ', canDelete, ', canEdit = ', canEdit )

            Ext.getCmp('userLoginInfo').setVisible( isSelf || authIsAdmin );
            configControl( Ext.getCmp('UserEditPrivilegeLevelCombo'), undefined, !authIsAdmin, undefined, undefined, undefined, undefined );

            configControl( Ext.getCmp('UserEditAgencyCombo'), undefined, !authIsAdmin, undefined, undefined, undefined, undefined );
            Ext.getCmp('user__is_auditor').setReadOnly(!authIsAdmin);
            Ext.getCmp('user__is_active').setReadOnly(!authIsAdmin);

            newBtn.setDisabled(!canCreate);
            copyBtn.setDisabled(!canCopy);
            deleteBtn.setDisabled(!canDelete);

            applyBtn.setDisabled(!canEdit);
            okBtn.setDisabled(!canEdit);

        }
        catch (ex){
            WxLog( "error", "controller/Users.enableAddEdit", ex.message );
        }

    }, //enableAddEdit

    filterPrivlegeLevelCombo: function() {

        try {

            var theCboStore = Ext.getCmp('UserEditPrivilegeLevelCombo').getStore();
            theCboStore.clearFilter(true);

            var privLevelId = 999999999999999;

            if( isDefined(document.app.userAuths) ) {

                privLevelId = parseInt(document.app.userAuths['userInfo'].privLevelId);

            }

            theCboStore.filter([
                {
                    filterFn: function (record) {
                        return ( record.get('id') >= privLevelId );
                    }
                }
            ]);

        }
        catch (ex){
            WxLog( "error", "controller/Users.filterPrivlegeLevelCombo", ex.message );
        }

    }, // filterPrivlegeLevelCombo
    
    filterAgencyCombo: function(){

        try {


            var theEditForm = Ext.getCmp('UserEditFormID').getForm();


            var theCboStore = Ext.getCmp('UserEditAgencyCombo').getStore();
            theCboStore.clearFilter(true);


            var theAuthUsersAgencies = [];
            var privLevelCode = 'none';

            if( isDefined(document.app.userAuths) ) {

                privLevelCode = document.app.userAuths['userInfo'].privLevelCode;
                var aCtr = 0;

                if (privLevelCode == 'site_admin') {

                    for (aCtr = 0; aCtr < theCboStore.getTotalCount(); aCtr++) {
                        theRecord = theCboStore.getAt(aCtr);
                        if( isDefined(theRecord) ){
                            theAuthUsersAgencies.push(parseInt(theRecord.data.id));
                        }
                    }


                } else {

                    for (aCtr = 0; aCtr < document.app.userAuths['userAgencies'].length; aCtr++) {
                        theAuthUsersAgencies.push(parseInt(document.app.userAuths['userAgencies'][aCtr].id));
                    }

                }

            }


            var theSelectedUserID = parseInt(theEditForm.findField('user__user_id').getValue());
            var theSelectedUsersAgencies = [];

            if( theSelectedUserID > 0 ) {

                var theSelectedAgencyID = parseInt(theEditForm.findField('user__agency_id').getValue());

                var theListStore = Ext.getCmp('UserListID').getStore();

                var rCtr = 0;
                var theRecord;

                for ( rCtr = 0; rCtr < theListStore.getTotalCount(); rCtr++ ) {

                    theRecord = theListStore.getAt(rCtr);

                    if( parseInt(theRecord.data.user__user_id) == theSelectedUserID ) {
                        theSelectedUsersAgencies.push(parseInt(theRecord.data.user__agency_id));
                    }

                }

            }

            var theSelectedPrivilegeID = parseInt(Ext.getCmp('UserEditPrivilegeLevelCombo').getValue());
            var minAgencyID = ( theSelectedPrivilegeID <= 2 ) ? -2 : 0;

            theCboStore.filter([
                {
                    filterFn: function (record) {

                        var retVal;

                        var thisAgencyID = record.get('id');

                        retVal = ( thisAgencyID >= minAgencyID ) &&
                            ( ( thisAgencyID == theSelectedAgencyID ) || (  ( theAuthUsersAgencies.indexOf(thisAgencyID) >= 0 ) &&
                                                                           !( theSelectedUsersAgencies.indexOf(thisAgencyID) >= 0 ) ) );

                        return retVal;

                    }

                }
            ]);

        }
        catch (ex){
            WxLog( "error", "controller/Users.filterAgencyCombo", ex.message );
        }

    }, //filterAgencyCombo


    userApply: function (hideForm) {

        var msg = 'Starting ...';
        WxLog("debug", "controller/Users.userApply", msg);

        try {

            hideForm = ( isDefined(hideForm) && ( hideForm == true || hideForm == false ) ) ? hideForm : false;


            var theEditForm = Ext.getCmp('UserEditFormID').getForm();


            var isNew = stringToBoolean(theEditForm.findField('user__is_new').getValue());


            var otherErrors = new Array();

            var newPassword = theEditForm.findField('user__new_password').getValue();

            if( isNew || ( newPassword != '' ) ) {

                var confirmPassword = theEditForm.findField('user__new_password_confirm').getValue();

                if( !isAcceptablePassword(newPassword) ){

                    otherErrors.push( document.app.passwordRulesDscr );

                } else if( confirmPassword != newPassword ){

                    otherErrors.push( "Passwords do not match." );

                } else {

                    newPassword = hex_md5(newPassword);

                }

            } else {
                newPassword = theEditForm.findField('user__user_password').getValue();
            }


            var userName = theEditForm.findField('user__user_name').getValue();

            if( isNew || ( userName !== Ext.getCmp('user__user_name').originalValue ) ) {

                var usernameIsUnique = stringToBoolean(theEditForm.findField('user__user_name_is_unique').getValue());

                if( !isAcceptableUsername(userName) ){

                    otherErrors.push( document.app.usernameRulesDscr );

                } else if( !usernameIsUnique) {

                    otherErrors.push( 'Username is not available.' );

                }

            }

            if( !theEditForm.isValid() || otherErrors.length > 0 ){
                showValidationErrors( 'UserEditFormID', otherErrors );
                return;
            }


            Ext.getCmp('userLoginInfo').setDisabled(false);
            Ext.getCmp('userPersonInfo').setDisabled(false);
            Ext.getCmp('userContactInfo').setDisabled(false);


            Ext.getCmp('UserEditFormID').getForm().submit({

                params: {
                    editing_user_id: document.app.user_id,
                    user__is_auditor: Ext.getCmp('user__is_auditor').getValue() ? 1 : 0,
                    user__is_active: Ext.getCmp('user__is_active').getValue() ? 1 : 0,
                    user__user_pword: newPassword
                },

                success: function(form, action) {

                    Ext.getBody().unmask();

                    WxLog("debug","controller/Users.userApply()","Saved User");

                    var updatedUserID = parseInt(action.result.data[0].mainID);

                    theEditForm.findField('user__selected_user_id').setValue(updatedUserID);


                    Ext.getStore('Users').load();


                    if( hideForm ) {
                        var thisCtrl = document.app.getController('Users');
                        thisCtrl.hide();
                    }

                },

                failure: function(form, action) {

                    WxLog( "error", "controller/Users.userApply", "Failed to save User" );

                    console.log(form);
                    console.log(action);

                    Ext.getBody().unmask();

                },

                exception: function(form, action) {
                    Ext.getBody().unmask();
                    WxLog( "error", "controller/Users.userApply", "Exception thrown while saving User" );
                }

            });

        }
        catch(ex) {
            WxLog("error", "controller/Users.userApply", ex.message);
        }
    }, //userApply

    uniqueUsernameResponse: function( isUnique ){

        if (isUnique) {
            Ext.getCmp('user__user_name_is_unique').setValue(true);
            Ext.getCmp('user__user_name_image').setSrc('images/12px-a_check.png');
        } else {
            Ext.getCmp('user__user_name_is_unique').setValue(false);
            Ext.getCmp('user__user_name_image').setSrc('images/12px-an_x.png');
            Ext.MessageBox.alert( 'Sorry ...', 'The username is not available.' )
        }

    },


    userNew: function(){

        var theEditForm = Ext.getCmp('UserEditFormID').getForm();

        var newBtn = Ext.getCmp('userNewButton');
        var copyBtn = Ext.getCmp('userCopyButton');
        var deleteBtn = Ext.getCmp('userDeleteButton');

        var theListGrid = Ext.getCmp('UserListID');
        theListGrid.getSelectionModel().deselectAll();

        this.clear();

        theEditForm.findField('user__is_new').setValue(true);

        this.filterAgencyCombo();

        newBtn.setDisabled(true);
        copyBtn.setDisabled(true);
        deleteBtn.setDisabled(true);

    }, //userNew



    userCopy: function(){

        var theEditForm = Ext.getCmp('UserEditFormID').getForm();

        var newBtn = Ext.getCmp('userNewButton');
        var copyBtn = Ext.getCmp('userCopyButton');
        var deleteBtn = Ext.getCmp('userDeleteButton');

        var theListGrid = Ext.getCmp('UserListID');
        theListGrid.getSelectionModel().deselectAll();

        theEditForm.findField('user__user_agency_id').setValue(null);
        theEditForm.findField('user__agency_id').setValue(null);
        Ext.getCmp('UserEditAgencyCombo').setValue(null);

        theEditForm.findField('user__is_copy').setValue(true);

        this.filterAgencyCombo();

        Ext.getCmp('userLoginInfo').setDisabled(true);
        Ext.getCmp('user__is_auditor').setReadOnly(true);
        Ext.getCmp('user__is_active').setReadOnly(true);
        Ext.getCmp('userPersonInfo').setDisabled(true);
        Ext.getCmp('userContactInfo').setDisabled(true);

        newBtn.setDisabled(true);
        copyBtn.setDisabled(true);
        deleteBtn.setDisabled(true);

    }, //userCopy



    userDelete: function () {

        var msg = 'Starting ...';
        WxLog("debug", "controller/Users.userDelete", msg);

        try {

            var theEditForm = Ext.getCmp('UserEditFormID').getForm();

            if( theEditForm.findField("user__user_id").getValue() == -1 ){

                Ext.MessageBox.alert("Information","There are no users to delete.");

                return;
            }

            var numAudits = theEditForm.findField('user__audit_count').getValue();

            if( numAudits > 0 ){

                var inUseText = "The user '" + Ext.getCmp('user__user_name').getValue() + "' cannot be deleted; it is associated with " + numAudits + " audit(s).</br>" +
                                "You can, instead, disable the user by unchecking 'Active'.";
                Ext.MessageBox.alert( "Information", inUseText );

                return;
            }

            var me = this;

            Ext.Msg.show({

                title: 'Confirm',
                msg: 'This will permenantly delete the user</br>'+
                     'Are you sure you want to continue?</br>',
                width: 300,
                icon: Ext.Msg.QUESTION,
                buttons: Ext.Msg.YESNO,
                scope: this,
                fn: this.userDeleteDialogResponse

            });

        }

        catch(ex) {
            WxLog("error", "controller/Users.userDelete", ex.message);
        }

    }, //userDelete

    userDeleteDialogResponse: function(btn, text) {

        console.log(btn);
        console.log(text);

        if(btn=="yes"){
            this.deleteUser()
        }

    }, // userDeleteDialogResponse

    deleteUser: function() {

        var msg = 'Starting ...';
        WxLog("debug", "controller/Users.deleteUser", msg);

        try {

            var theUserID = Ext.getCmp('UserEditFormID').getForm().findField("user__user_id").getValue();
            var theAgencyID = Ext.getCmp('UserEditFormID').getForm().findField("user__agency_id").getValue();
            msg = 'Delete started, user ID = ' + theUserID + ', the agency ID = ' + theAgencyID;
            WxLog("debug", "controller/Users.deleteUser", msg);

            var theURL =  'data/User.php?command=delete';

            Ext.Ajax.request({

                url: theURL,

                params: {
                    editing_user_id: document.app.user_id,
                    user__user_id: theUserID,
                    user__agency_id: theAgencyID
                },

                success: function(form, action) {

                    var theForm = form;
                    var theAction = action;

                    document.app.getController('Users').deleteCompleted(theForm, theAction);

                    msg = 'Finished';
                    WxLog("debug", "controller/Users.deleteUser", msg);

                },

                failure: function(form, action) {

                    msg = 'Failure';
                    WxLog("error", "controller/Users.deleteUser", msg);

                    Ext.MessageBox.confirm('Confirm', 'there was a problem communicating with the server, do you wish to retry?', this.deleteUser);

                },

                exception: httpExceptionHandler

            });

        } //try

        catch(ex) {
            WxLog("error", "controller/Users.deleteUser", ex.message);
        }

    }, //deleteUser

    deleteCompleted: function(form, action) {

        var msg = 'Starting ...';
        WxLog("debug", "controller/MultiUsers.deleteCompleted", msg);

        try {

            this.clear();

            Ext.getStore('Users').load();

            msg = 'Finished';
            WxLog("debug", "controller/Users.deleteCompleted", msg);

        }
        catch(ex) {
            WxLog("error", "controller/Users.deleteCompleted", ex.message);
        }

    }, //deleteCompleted




    resumeComponentEvents: function() {
        resumeComponentEvents('Users', 'UserEditID', 'UserListID', 'user_edit_toolbar');
    } //resumeComponentEvents


});