Ext.define('wx.model.User', {
    extend: 'Ext.data.Model',
    idProperty: 'user__unique_id',
    fields: [

        { name: 'user__unique_id', type: 'int', defaultValue: -1 },


        { name: 'user__user_id', type: 'int', defaultValue: -1, useNull: true },
        { name: 'user__user_name', defaultValue: null, useNull: true },
        { name: 'user__user_name_is_unique', type: 'boolean', defaultValue: null, useNull: true },
        { name: 'user__user_password', defaultValue: null, useNull: true },
        { name: 'user__auth_group_id', type: 'int', defaultValue: null, useNull: true },
        { name: 'user__auth_group', defaultValue: null, useNull: true },
        { name: 'user__user_comments', defaultValue: null, useNull: true },
        { name: 'user__user_modified', type: 'date', dateFormat: 'Y-m-d H:i:s' },


        { name: 'user__user_agency_id', type: 'int', defaultValue: null, useNull: true },
        { name: 'user__agency_id', type: 'int', defaultValue: null, useNull: true },
        { name: 'user__agency', defaultValue: null, useNull: true },
        { name: 'user__is_auditor', type: 'boolean', defaultValue: null, useNull: true },
        { name: 'user__is_active', type: 'boolean', defaultValue: null, useNull: true },
        { name: 'user__audit_count', type: 'int', defaultValue: 0 },


        { name: 'user__contact_id', type: 'int', defaultValue: -1, useNull: true },
        { name: 'user__long_name', defaultValue: null, useNull: true },
        { name: 'user__prefix_id', type: 'int', defaultValue: null, useNull: true },
        { name: 'user__prefix', defaultValue: null, useNull: true },
        { name: 'user__first_name', defaultValue: null, useNull: true },
        { name: 'user__middle_initial', defaultValue: null, useNull: true },
        { name: 'user__last_name', defaultValue: null, useNull: true },
        { name: 'user__suffix_id', type: 'int', defaultValue: null, useNull: true },
        { name: 'user__suffix', defaultValue: null, useNull: true },
        { name: 'user__company', defaultValue: null, useNull: true },
        { name: 'user__title', defaultValue: null, useNull: true },

        { name: 'user__street_address', defaultValue: null, useNull: true },
        { name: 'user__unit_number', defaultValue: null, useNull: true },
        { name: 'user__city', defaultValue: null, useNull: true },
        { name: 'user__state_id', type: 'int', defaultValue: null, useNull: true },
        { name: 'user__state', defaultValue: null, useNull: true },
        { name: 'user__zip_code', defaultValue: null, useNull: true },

        { name: 'user__work_phone', defaultValue: null, useNull: true },
        { name: 'user__cell_phone', defaultValue: null, useNull: true },
        { name: 'user__home_phone', defaultValue: null, useNull: true },
        { name: 'user__fax_phone', defaultValue: null, useNull: true },

        { name: 'user__email_address', defaultValue: null, useNull: true },
        { name: 'user__web_address', defaultValue: null, useNull: true },

        { name: 'user__contact_comments', defaultValue: null, useNull: true },
        { name: 'user__contact_modified', type: 'date', dateFormat: 'Y-m-d H:i:s' },

        { name: 'user__is_new', type: 'boolean', defaultValue: false, useNull: true },
        { name: 'user__is_copy', type: 'boolean', defaultValue: false, useNull: true }

    ]
});

