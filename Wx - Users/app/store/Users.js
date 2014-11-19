Ext.define('wx.store.Users', {
    extend: 'Ext.data.Store',
    model: 'wx.model.User',
    autoLoad: false,
    alias: 'widget.userstore',

    proxy: {

        type: 'ajax',

        api: {
            read: 'data/User.php?command=list',
            update: 'data/User.php?command=update'
        },

        reader: {
            type: 'json',
            root: 'data',
            successProperty: 'success'
        },

        actionMethods: {
            read: 'POST',
            update: 'POST'
        }

    },

    listeners: {

        beforeload: function(str, oper, opts) {

            var app = document.app;
            str.setExtraParam( 'editing_user_id', document.app.user_id );

        },

        load: function(str, records, success, operation, opts){

            wx.app.getController('Users').onUsersStoreLoad(this);

        }

    }

});
