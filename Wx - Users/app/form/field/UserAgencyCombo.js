Ext.define('wx.form.field.UserAgencyCombo', {
    requires: [ 'Ext.util.Sortable', 'wx.model.UserAgencyCombo' ],
    extend: 'Ext.form.field.ComboBox',
    mixins: { sortable: 'Ext.util.Sortable' },
    alias: 'widget.useragencycombo',

    constructor: function(config) {

        this.superclass.constructor.apply( this, new Array(config) );

        this.sprocname = null;
        this.sprocmodel = null;
        this.allowBlank = false;

        if ( config ) {
            if ( config.sprocname ) {
                this.sprocname = config.sprocname;
            }
            if( config.requiredItem ){
                this.emptyText = config.requiredItem;
            }
            if ( config.theValueField ) {
                this.name = config.theValueField;
            }
            if ( config.allowBlank ) {
                this.allowBlank = config.allowBlank;
            }
            if(config.controllerName){
                this.controllerName = config.controllerName;
            }
        }

        this.store.on( 'load', function ( store, records, successful, options ) {
            this.getPicker().setLoading(false);
        }, this );

    },

    initComponent: function() {

        Ext.apply(this, {

            store: Ext.create('Ext.data.Store', {

                storeId: this.storeId,
                parentComboId: this.id,
                theControllerName: this.controllerName,
                model: 'wx.model.UserAgencyCombo',
                autoLoad: false,

                proxy: {
                    type: 'ajax',
                    noCache: true,
                    pageParam: undefined,
                    limitParam: undefined,
                    startParam: undefined,
                    url: 'data/getUserAgencies.php?sproc=' + this.sprocname,
                    method: 'GET',
                    reader: Ext.create( 'Ext.data.reader.Json', {
                        type: 'json',
                        root: 'data'
                    })
                },

                listeners: {

                    beforeload: function(str, oper, opts) {
                        //console.log('form/field/UserAgencyCombo.store.beforeLoad(',this.theControllerName,') event');
                        var app = document.app;
                        str.setExtraParam('user_id', app.user_id);
                        str.setExtraParam('audit_type_id', app.audit_type_id);

                    },

                    load: function(str, records, success, operation, opts) {
                        //console.log('form/field/UserAgencyCombo.store.load(',this.theControllerName,') event');

                        if( success ) {

                            if( this.theControllerName ){
                                document.app.getController(this.theControllerName).initialLoadCall();
                                //document.app.getController(this.theControllerName).filterUserAgencyCombo(str.parentComboId);
                            }

                        } else {
                            Ext.Msg.alert('*****************       Failed to load the agency list.');
                        }

                    }
                }

            }),

            model: 'wx.model.UserAgencyCombo',
            queryMode: "local",
            buffered: false,
            triggerAction: "all",
            displayField: this.displayField,
            valueField: 'id',
            editable: false

        });

        this.callParent(arguments);

        this.store.proxy.reader.setModel(this.store.model);

    }

});
