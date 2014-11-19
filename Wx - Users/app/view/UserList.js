Ext.define('wx.view.common.UserList', {
    extend: 'Ext.grid.Panel',
    require: [ 'Ext.ux.FilterBar.*' ],
    alias : 'widget.userlist',
    title : 'All Users',
    plugins: [{
        ptype: 'filterbar',
        renderHidden: true,
        showShowHideButton: true,
        showClearAllButton: true
    }],
    store  : 'Users',
    height: 215,
    margin: '5 0 0 0',
    autoShow: true,
    autoRender: 'wxworkspacefooter',
    id: 'UserListID',

    initComponent: function () {

        this.enableBubble('load');
        this.callParent(arguments);

    },

    columns: {

        plugins: [{
            ptype: 'gridautoresizer'
        }],
        items: [

            { header: 'Name', dataIndex: 'user__long_name', width: 250, filter: 'combo' },
            { header: 'Username', dataIndex: 'user__user_name', width: 100, filter: 'combo' },
            { header: 'Agency', dataIndex: 'user__agency', width: 275, filter: 'combo' },
            { header: 'Privilege Level', dataIndex: 'user__auth_group', width: 150, filter: 'combo' },
            { header: 'Is Active', dataIndex: 'user__is_active', width: 60, renderer: function(value){ return ( value == 1 ) ? 'Yes' : 'No'; }, filter: true },
            { header: 'Num Audits', dataIndex: 'user__audit_count', width: 80, filter: true },
            { header: 'Last Edited', dataIndex: 'user__user_modified', width: 100, flex: 1, renderer: Ext.util.Format.dateRenderer(document.app.date_time_format), filter: true }

        ]

    },

    listeners: {

        select: function( theGrid, theRecord, theRecordIndex ){
            wx.app.getController('Users').onUserListChange(theRecord);
        }

    }

});
