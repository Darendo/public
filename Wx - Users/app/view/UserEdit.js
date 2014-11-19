Ext.define('wx.view.common.UserEdit', {
    extend: 'Ext.form.Panel',
    requires: [ 'Ext.form.*',
                'Ext.grid.*',
                'wx.form.field.PasswordMeter',
                'wx.form.field.UserAgencyCombo',
                'Ext.form.ComboBox',
                'wx.form.field.EnumCombo' ],
    alias: 'widget.useredit',
    title: 'User Details',
    autoShow: true,
    autoRender: 'wxworkspacecenter',
    height: 621,
    id: 'UserEditID',

    initComponent: function() {

        this.callParent(arguments);

    },

    items: [
        {
            xtype: 'form',
            id: 'UserEditFormID',
            autoWidth: true,
            trackResetOnLoad: true,
            border:false,
            autoHeight: false,
            autoScroll: true,

            url: 'data/User.php?command=update',
            baseParams: {
                audit_type_id: document.app.audit_type_id,
                user_id: document.app.user_id
            },

            items: [

                new Ext.form.Hidden({
                    id: 'user__user_id',
                    name: 'user__user_id'
                }),
                new Ext.form.Hidden({
                    id: 'user__user_password',
                    name: 'user__user_password'
                }),

                new Ext.form.Hidden({
                    id: 'user__user_agency_id',
                    name: 'user__user_agency_id'
                }),
                new Ext.form.Hidden({
                    id: 'user__contact_id',
                    name: 'user__contact_id'
                }),
                new Ext.form.Hidden({
                    id: 'user__audit_count',
                    name: 'user__audit_count'
                }),

                new Ext.form.Hidden({
                    xtype: 'numberfield',
                    id: 'user__selected_user_id',
                    name: 'user__selected_user_id'
                }),
                new Ext.form.Hidden({
                    id: 'user__selected_user_name',
                    name: 'user__selected_user_name'
                }),
                new Ext.form.Hidden({
                    id: 'user__user_name_is_unique',
                    name: 'user__user_name_is_unique'
                }),

                new Ext.form.Hidden({
                    xtype: 'numberfield',
                    id: 'user__is_copy',
                    name: 'user__is_copy'
                }),
                new Ext.form.Hidden({
                    xtype: 'numberfield',
                    id: 'user__is_new',
                    name: 'user__is_new'
                }),

                {
                    xtype: 'panel',
                    height: 565,
                    border:false,
                    autoScroll: true,
                    autoHeight: false,

                    items: [

                        {
                            border: false,
                            items: [

                                {
                                    layout:'column',
                                    border:false,
                                    items: [
                                        {
                                            columnWidth: 1/2,
                                            border: false,
                                            style : 'margin-top:10px',
                                            layout: 'anchor',
                                            //hidden: true,
                                            items:[

                                                {

                                                    xtype: 'fieldset',
                                                    title: 'Login Information',
                                                    id: 'userLoginInfo',
                                                    name: 'userLoginInfo',
                                                    collapsible: false,
                                                    autoHeight: true,
                                                    collapsed: false,
                                                    style: 'margin-left: 10px; margin-right: 10px; ',
                                                    layout: 'anchor',

                                                    items: [

                                                        {
                                                            xtype: 'fieldcontainer',
                                                            border: false,
                                                            layout: 'hbox',
                                                            width: '100%',

                                                            items: [

                                                                {
                                                                    xtype: 'textfield',
                                                                    id: 'user__user_name',
                                                                    name: 'user__user_name',
                                                                    fieldLabel: 'Username',
                                                                    allowBlank: false,
                                                                    minLength: 1,
                                                                    minLengthText: 'The username must be between 1 and 50 characters.',
                                                                    maxLength: 50,
                                                                    maxLengthText: 'The username must be between 1 and 50 characters.', //'The maximum length for this field is 50 characters.',
                                                                    style: 'margin-left:20px; margin-right:10px;',
                                                                    listeners: {
                                                                        blur: function( theObj, theEventObj, eOpts ){

                                                                            try{

                                                                                if( isDefined(theObj) ){

                                                                                    var newValue = theObj.getValue();

                                                                                    if( isDefined(newValue) ){

                                                                                        if( newValue !== theObj.originalValue ){

                                                                                            Ext.getCmp('user__selected_user_name').setValue(newValue);

                                                                                            Ext.getCmp('user__user_name_is_unique').setValue(null);

                                                                                            var isValidUsername = isAcceptableUsername(newValue);

                                                                                            if (isValidUsername) {
                                                                                                isUniqueUsername(newValue);
                                                                                            } else {
                                                                                                Ext.getCmp('user__user_name_image').setSrc('images/12px-an_x.png');
                                                                                            }

                                                                                        } else {
                                                                                            Ext.getCmp('user__user_name_is_unique').setValue( Ext.getCmp('user__user_name_is_unique').originalValue );
                                                                                            Ext.getCmp('user__user_name_image').setSrc('');
                                                                                        }

                                                                                    }
                                                                                }
                                                                            }
                                                                            catch (ex){
                                                                                WxLog( "error", "view/UserEdit.user__user_name.blue()", ex.message );
                                                                            }
                                                                        }
                                                                    }

                                                                },

                                                                Ext.create('Ext.Img', {
                                                                        id: 'user__user_name_image',
                                                                        name: 'user__user_name_image',
                                                                        width: 12,
                                                                        height: 12,
                                                                        style: 'margin-top:5px;'
                                                                    }
                                                                )

                                                            ]

                                                        },

                                                        {
                                                            xtype: 'fieldcontainer',
                                                            border: false,
                                                            layout: 'hbox',
                                                            width: '100%',

                                                            items: [

                                                                Ext.create('wx.form.field.PasswordMeter', {
                                                                    id: 'user__new_password',
                                                                    name: 'user__new_password',
                                                                    fieldLabel: 'Password',
                                                                    allowBlank: true,
                                                                    maxLength: 64,
                                                                    inputType: 'password',
                                                                    enableKeyEvents: true,
                                                                    fieldValidationErrorLabel: 'Password must be at least 8 characters and have at least 1 of each of the following: capital letters, small letters, numbers, and punctuation.',
                                                                    maxLengthText: 'The maximum length for this field is 64 characters.',
                                                                    //anchor: '100%',
                                                                    style: 'margin-left:20px; margin-right:10px;',
                                                                    width: '60%',
                                                                    validationEvent: "blur",
                                                                    // This custom validator option expects a return value of boolean true if it
                                                                    // validates, and a string with an error message if it doesn't
                                                                    validator: function () {

                                                                        try {

                                                                            var editForm = Ext.getCmp("UserEditFormID").getForm();

                                                                            var pass1 = editForm.findField('user__new_password').getValue();
                                                                            var pass2 = editForm.findField('user__new_password_confirm').getValue();

                                                                            if (isDefined(pass1) && pass1 != '' && isDefined(pass2) && pass2 != '') {

                                                                                if (pass1 == pass2) {
                                                                                    return true;
                                                                                } else {
                                                                                    return "Passwords do not match!";
                                                                                }

                                                                            } else {
                                                                                return true;
                                                                            }
                                                                        }
                                                                        catch (ex){
                                                                            WxLog( "error", "view/UserEdit.user__new_password.validator()", ex.message );
                                                                            return false;
                                                                        }

                                                                    },
                                                                    listeners: {

                                                                        render: function(c) {
                                                                            Ext.create('Ext.tip.ToolTip', {
                                                                                target: c.getEl(),
                                                                                html: document.app.passwordRulesDscr,
                                                                                dismissDelay: 0
                                                                            });
                                                                        },

                                                                        keyup: function (fld, e, eOpts) {

                                                                            try {


                                                                                if( fld.id == 'user__new_password'){

                                                                                    var pwdFld = Ext.getCmp('user__new_password');

                                                                                    if (isDefined(pwdFld)) {

                                                                                        var password = pwdFld.getValue();

                                                                                        if( password ){

                                                                                            var isOK = isAcceptablePassword(pwdFld.getValue());

                                                                                            if( isOK ) {
                                                                                                Ext.getCmp('user__new_password_image').setSrc('images/12px-a_check.png');
                                                                                            } else {
                                                                                                Ext.getCmp('user__new_password_image').setSrc('images/12px-an_x.png');
                                                                                            }

                                                                                        } else {
                                                                                            Ext.getCmp('user__new_password_image').setSrc('');
                                                                                        }
                                                                                    }

                                                                                }

                                                                            }
                                                                            catch (ex){
                                                                                WxLog( "error", "view/UserEdit.user__new_password.keyup()", ex.message );
                                                                            }

                                                                        }
                                                                    }

                                                                }),

                                                                Ext.create('Ext.Img', {
                                                                        id: 'user__new_password_image',
                                                                        name: 'user__new_password_image',
                                                                        width: 12,
                                                                        height: 12,
                                                                        style: 'margin-top:5px;'
                                                                    }
                                                                )
                                                            ]

                                                        },


                                                        {
                                                            xtype: 'fieldcontainer',
                                                            border: false,
                                                            layout: 'hbox',
                                                            width: '100%',

                                                            items: [



                                                                {
                                                                    xtype: 'textfield',
                                                                    id: 'user__new_password_confirm',
                                                                    name: 'user__new_password_confirm',
                                                                    fieldLabel: 'Confirm Password',
                                                                    allowBlank: true,
                                                                    maxLength: 64,
                                                                    inputType: 'password',
                                                                    enableKeyEvents: true,
                                                                    maxLengthText: 'The maximum length for this field is 64 characters.',
                                                                    style: 'margin-left:20px; margin-right:10px;',
                                                                    width: '60%',
                                                                    listeners: {

                                                                        keyup: function (fld, e, eOpts) {

                                                                            try {

                                                                                if( fld.id == 'user__new_password_confirm'){

                                                                                    var editForm = Ext.getCmp("UserEditFormID").getForm();
                                                                                    var pass1 = editForm.findField('user__new_password').getValue();
                                                                                    var pass2 = editForm.findField('user__new_password_confirm').getValue();

                                                                                    if ( isDefined(pass1) && pass1 != '' && isDefined(pass2) && pass2 != '' ){

                                                                                        if (pass1 == pass2) {
                                                                                            Ext.getCmp('user__new_password_confirm_image').setSrc('images/12px-a_check.png');
                                                                                            return true;
                                                                                        } else {
                                                                                            Ext.getCmp('user__new_password_confirm_image').setSrc('images/12px-an_x.png');
                                                                                            return "Passwords do not match!";
                                                                                        }

                                                                                    } else {
                                                                                        Ext.getCmp('user__new_password_confirm_image').setSrc('');
                                                                                        return true;
                                                                                    }

                                                                                }

                                                                            }
                                                                            catch (ex){
                                                                                WxLog( "error", "view/UserEdit.user__new_password.keyup()", ex.message );
                                                                            }

                                                                        }
                                                                    }
                                                                },

                                                                Ext.create('Ext.Img', {
                                                                        id: 'user__new_password_confirm_image',
                                                                        name: 'user__new_password_confirm_image',
                                                                        width: 12,
                                                                        height: 12,
                                                                        style: 'margin-top:5px;'
                                                                    }
                                                                )
                                                            ]

                                                        },

                                                        Ext.create('wx.form.field.EnumCombo', {
                                                            id: 'UserEditPrivilegeLevelCombo',
                                                            fieldLabel: 'Privilege Level',
                                                            controllerName: 'Users',
                                                            autoLoad: false,
                                                            inputId: 'user__auth_group_id',
                                                            storeId: 'e_auth_group',
                                                            allowBlank: false,
                                                            theValueField: 'user__auth_group_id',
                                                            theTextField: 'user__auth_group',
                                                            style : 'margin-left:20px; margin-right:10px;',
                                                            lastQuery: '',
                                                            listeners: {

                                                                change: function ( theCbo, newValue, oldValue, eOpts ) {

                                                                    var theAgencyCbo = Ext.getCmp('UserEditAgencyCombo');
                                                                    var theRawValue = theAgencyCbo.getRawValue();

                                                                    theAgencyCbo.store.clearFilter(true);
                                                                    var isEnabled = true;

                                                                    switch(newValue){
                                                                        case 1:

                                                                            isEnabled = false;
                                                                            theRawValue = 'All Agencies';

                                                                            break;


                                                                        case 2:

                                                                            isEnabled = false;
                                                                            theRawValue = 'All Agencies in Same-state';

                                                                            break;


                                                                        default:

                                                                            // Leave the vars as-is

                                                                            theAgencyCbo.store.filter([{
                                                                                filterFn: function (record) {
                                                                                    var showRec = ( record.get('id') > 0 );
                                                                                    return showRec;
                                                                                }
                                                                            }]);

                                                                            break;

                                                                    }

                                                                    theAgencyCbo.setRawValue(theRawValue);
                                                                    configControl( theAgencyCbo, isEnabled, !isEnabled, undefined, undefined, undefined, !isEnabled, true );

                                                                }
                                                            }
                                                        })

                                                    ]

                                                },

                                                {

                                                    xtype: 'fieldset',
                                                    title: 'Agency Info',
                                                    id: 'UserEditAgencyInfo',
                                                    name: 'UserEditAgencyInfo',
                                                    collapsible: false,
                                                    autoHeight: true,
                                                    collapsed: false,
                                                    style: 'margin-left: 10px; margin-right: 10px; ',
                                                    layout: 'anchor',

                                                    items: [

                                                        Ext.create('wx.form.field.UserAgencyCombo', {
                                                            id: 'UserEditAgencyCombo',
                                                            controllerName: 'Users',
                                                            autoLoad: false,
                                                            allowBlank: false,
                                                            sprocname: 'user_agency_list_select',
                                                            fieldLabel: 'Agency',
                                                            style: 'margin-left: 20px;',
                                                            displayField: 'agency_name',
                                                            theValueField: 'user__agency_id',
                                                            theTextField: 'user__agency',
                                                            width: 460,
                                                            lastQuery: ''
                                                        }),

                                                        {
                                                            xtype: 'checkbox',
                                                            id: 'user__is_auditor',
                                                            name: 'user__is_auditor',
                                                            fieldLabel: 'Auditor',
                                                            style : 'margin-left:20px; margin-right:10px;'
                                                        },

                                                        {
                                                            xtype: 'checkbox',
                                                            id: 'user__is_active',
                                                            name: 'user__is_active',
                                                            fieldLabel: 'Active',
                                                            style : 'margin-left:20px; margin-right:10px;'
                                                        }

                                                    ]

                                                },

                                                {

                                                    xtype: 'fieldset',
                                                    title: 'Person Information',
                                                    id: 'userPersonInfo',
                                                    name: 'userPersonInfo',
                                                    collapsible: false,
                                                    autoHeight: true,
                                                    collapsed: false,
                                                    style: 'margin-left: 10px; margin-right: 10px; ',
                                                    layout: 'anchor',

                                                    items: [

                                                        Ext.create('wx.form.field.EnumCombo', {
                                                            id: 'UserEditUserPrefixCombo',
                                                            fieldLabel: 'Prefix',
                                                            controllerName: 'Users',
                                                            autoLoad: false,
                                                            inputId: 'user__prefix_id',
                                                            storeId: 'e_name_prefix',
                                                            allowBlank: true,
                                                            theValueField: 'user__prefix_id',
                                                            theTextField: 'user__prefix',
                                                            style : 'margin-left:20px; margin-right:10px;',
                                                            width: 175
                                                        }),

                                                        {
                                                            xtype: 'textfield',
                                                            fieldLabel: 'First',
                                                            id: 'user__first_name',
                                                            name: 'user__first_name',
                                                            allowBlank: true,
                                                            maxLength: 90,
                                                            maxLengthText: 'The maximum length for this field is 90 characters.',
                                                            style: 'margin-left:20px; margin-right:10px;',
                                                            width: 250
                                                        },

                                                        {
                                                            xtype: 'textfield',
                                                            fieldLabel: 'Middle Init',
                                                            id: 'user__middle_initial',
                                                            name: 'user__middle_initial',
                                                            allowBlank: true,
                                                            maxLength: 1,
                                                            maxLengthText: 'The maximum length for this field is 1 character.',
                                                            style: 'margin-left:20px; margin-right:10px;',
                                                            width: 125
                                                        },

                                                        Ext.create('wx.form.field.EnumCombo', {
                                                            id: 'UserEditUserSuffixCombo',
                                                            fieldLabel: 'Suffix',
                                                            controllerName: 'Users',
                                                            autoLoad: false,
                                                            inputId: 'user__suffix_id',
                                                            storeId: 'e_name_suffix',
                                                            allowBlank: true,
                                                            theValueField: 'user__suffix_id',
                                                            theTextField: 'user__suffix',
                                                            style : 'margin-left:20px; margin-right:10px;',
                                                            width: 175
                                                        }),

                                                        {
                                                            xtype: 'textfield',
                                                            fieldLabel: 'Last',
                                                            id: 'user__last_name',
                                                            name: 'user__last_name',
                                                            allowBlank: false,
                                                            maxLength: 120,
                                                            maxLengthText: 'The maximum length for this field is 120 characters.',
                                                            style: 'margin-left:20px; margin-right:10px;',
                                                            width: 400
                                                        },

                                                        {
                                                            xtype: 'textfield',
                                                            fieldLabel: 'Company',
                                                            id: 'user__company',
                                                            name: 'user__company',
                                                            allowBlank: true,
                                                            maxLength: 64,
                                                            maxLengthText: 'The maximum length for this field is 64 characters.',
                                                            style: 'margin-left:20px; margin-right:10px;',
                                                            width: 450
                                                        },

                                                        {
                                                            xtype: 'textfield',
                                                            fieldLabel: 'Title',
                                                            id: 'user__title',
                                                            name: 'user__title',
                                                            allowBlank: true,
                                                            maxLength: 64,
                                                            maxLengthText: 'The maximum length for this field is 64 characters.',
                                                            style: 'margin-left:20px; margin-right:10px;',
                                                            width: 450
                                                        }

                                                    ]
                                                }
                                            ]
                                        },

                                        {

                                            columnWidth: 1/2,
                                            border: false,
                                            style : 'margin-top:10px;',
                                            layout: 'anchor',
                                            items:[

                                                {

                                                    xtype: 'fieldset',
                                                    title: 'Contact Information',
                                                    id: 'userContactInfo',
                                                    name: 'userContactInfo',
                                                    collapsible: false,
                                                    autoHeight: true,
                                                    collapsed: false,
                                                    style: 'margin-left: 10px; margin-right: 10px; ',
                                                    layout: 'anchor',

                                                    items: [

                                                        {
                                                            xtype: 'textfield',
                                                            fieldLabel: 'Address',
                                                            name: 'user__street_address',
                                                            id: 'user__street_address',
                                                            maxLength: 255,
                                                            maxLengthText: 'The maximum length for this field is 255 characters.',
                                                            anchor: '100%',
                                                            style: 'margin-left:20px; margin-right:10px;'
                                                        },

                                                        {
                                                            xtype: 'textfield',
                                                            fieldLabel: 'Unit Number',
                                                            name: 'user__unit_number',
                                                            id: 'user__unit_number',
                                                            maxLength: 45,
                                                            maxLengthText: 'The maximum length for this field is 45 characters.',
                                                            style: 'margin-left:20px; margin-right:10px;',
                                                            width: 250
                                                        },

                                                        {
                                                            xtype: 'textfield',
                                                            fieldLabel: 'City',
                                                            id: 'user__city',
                                                            name: 'user__city',
                                                            maxLength: 80,
                                                            maxLengthText: 'The maximum length for this field is 80 characters.',
                                                            style: 'margin-left:20px; margin-right:10px;',
                                                            width: 400
                                                        },

                                                        Ext.create('wx.form.field.StateCombo', {
                                                            id: 'UserEditStateCombo',
                                                            controllerName: 'Users',
                                                            autoLoad: false,
                                                            inputId: 'user__state_id',
                                                            storeId: 'e_state',
                                                            allowBlank: false,
                                                            fieldLabel: 'State',
                                                            theValueField: 'user__state_id',
                                                            theTextField: 'user__state',
                                                            style: 'margin-left:20px; margin-right:10px;',
                                                            width: 150
                                                        }),

                                                        {
                                                            xtype: 'textfield',
                                                            vtype: 'zipCode',
                                                            fieldLabel: 'Zip code',
                                                            id: 'user__zip_code',
                                                            name: 'user__zip_code',
                                                            maxLength: 20,
                                                            maxLengthText: 'The maximum length for this field is 20 characters.',
                                                            style: 'margin-left:20px; margin-right:10px;',
                                                            width: 200
                                                        },


                                                        {
                                                            xtype: 'textfield',
                                                            vtype: 'phone',
                                                            fieldLabel: 'Work Phone',
                                                            id: 'user__work_phone',
                                                            name: 'user__work_phone',
                                                            maxLength: 20,
                                                            maxLengthText: 'The maximum length for this field is 20 characters.',
                                                            style: 'margin-left:20px; margin-right:10px; margin-top:25px;'
                                                        },

                                                        {
                                                            xtype: 'textfield',
                                                            vtype: 'phone',
                                                            fieldLabel: 'Cell Phone',
                                                            id: 'user__cell_phone',
                                                            name: 'user__cell_phone',
                                                            maxLength: 20,
                                                            maxLengthText: 'The maximum length for this field is 20 characters.',
                                                            style: 'margin-left:20px; margin-right:10px;'
                                                        },

                                                        {
                                                            xtype: 'textfield',
                                                            vtype: 'phone',
                                                            fieldLabel: 'Home Phone',
                                                            id: 'user__home_phone',
                                                            name: 'user__home_phone',
                                                            maxLength: 20,
                                                            maxLengthText: 'The maximum length for this field is 20 characters.',
                                                            style: 'margin-left:20px; margin-right:10px;'
                                                        },

                                                        {
                                                            xtype: 'textfield',
                                                            vtype: 'fax',
                                                            fieldLabel: 'Fax',
                                                            id: 'user__fax_phone',
                                                            name: 'user__fax_phone',
                                                            maxLength: 20,
                                                            maxLengthText: 'The maximum length for this field is 20 characters.',
                                                            style: 'margin-left:20px; margin-right:10px;'
                                                        },

                                                        {
                                                            xtype: 'textfield',
                                                            vtype: 'email',
                                                            fieldLabel: 'Email',
                                                            id: 'user__email_address',
                                                            name: 'user__email_address',
                                                            anchor: '100%',
                                                            maxLength: 255,
                                                            maxLengthText: 'The maximum length for this field is 255 characters.',
                                                            style: 'margin-left:20px; margin-right:10px; margin-top:25px;'
                                                        },

                                                        {
                                                            xtype: 'textfield',
                                                            vtype: 'url',
                                                            fieldLabel: 'Web Page URL',
                                                            id: 'user__web_address',
                                                            name: 'user__web_address',
                                                            anchor: '100%',
                                                            maxLength: 255,
                                                            maxLengthText: 'The maximum length for this field is 255 characters.',
                                                            style: 'margin-left:20px; margin-right:10px;'
                                                        }

                                                    ]

                                                }

                                            ]

                                        }
                                    ]
                                },


                                {
                                    xtype:'fieldset',
                                    title: 'Comments',
                                    collapsible: false,
                                    autoHeight:true,
                                    collapsed: false,
                                    style : 'margin-left: 10px; margin-right: 10px; ',
                                    layout:'anchor',
                                    items: [
                                        {
                                            style : 'margin-top: 5px; margin-bottom: 10px; ',
                                            xtype:'textareafield',
                                            name:'user__comments',
                                            height: 100,
                                            anchor:'100%'
                                        }
                                    ]


                                }
                            ]
                        }
                    ]
                },

                Ext.create('Ext.toolbar.Toolbar', {
                    id: 'user_edit_toolbar',
                    bubbleEvents: ['formNew', 'formCopy', 'formDelete', 'formOk', 'formApply', 'formCancel'],
                    border: false,
                    cls: 'wx-f',
                    height: 31,
                    items: [
                        {
                            id: 'userNewButton',
                            name: 'userNewButton',
                            text: 'New',
                            disabled: true,
                            action: 'formNew',
                            cls: 'wx-view-btn'
                        },
                        {
                            id: 'userCopyButton',
                            name: 'userCopyButton',
                            text: 'Assign to Additional Agency',
                            disabled: true,
                            action: 'formCopy',
                            cls: 'wx-view-btn'
                        },
                        {
                            id: 'userDeleteButton',
                            name: 'userDeleteButton',
                            text: 'Delete',
                            disabled: true,
                            action: 'formDelete',
                            cls: 'wx-view-btn'
                        },
                        '->',
                        {
                            id: 'userOkButton',
                            name: 'userOkButton',
                            text: 'OK',
                            disabled: true,
                            action: 'formOk',
                            cls: 'wx-view-btn'
                        },
                        {
                            id: 'userApplyButton',
                            name: 'userApplyButton',
                            text: 'Apply',
                            disabled: true,
                            action: 'formApply',
                            cls: 'wx-view-btn'
                        },
                        {
                            text: 'Cancel',
                            action: 'formCancel',
                            cls: 'wx-view-btn'
                        }
                    ]
                })
            ]
        }
    ]
});
