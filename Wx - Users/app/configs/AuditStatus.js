Ext.define('wx.configs.AuditStatus', {
    alias: 'widget.auditstatus',

    singleton: true,

    auditID: null,
    auditJobID: null,
    auditTypeID: null,

    editStatus: null,
    runStatus: null,
    statusMsg: null,

    requestSubmitted: false,

    constructor: function(options) {

        this.initConfig(options);
        return this;

    },

    wxWsURL: function(){

        var tomcat_protocol = 'http';
        var tomcat_server = web_server.config.server_name;
        var tomcat_port = '8080';
        var ws_root = 'wx_ws_client';

        return tomcat_protocol+'://'+tomcat_server+':'+tomcat_port+'/'+ws_root;

    },

    loadAuditStatus: function ( audit_id, refresh_edit_status ) {

        this.auditID = audit_id;

        var theEntryPoint = 'RetrieveAuditRunStatus';

        var refreshEditStatus = ( isDefined(refresh_edit_status) ) ? refresh_edit_status : false;

        var theURL = this.wxWsURL() + '/' + theEntryPoint + '?' +
                    'audit_id=' + this.auditID + '&' +
                    'refresh_edit_status=' + refreshEditStatus;

        this.submitRequest( theURL );

    },

    refreshAuditStatus: function ( audit_id ) {

        try {

            this.requestSubmitted = true;

            this.auditID = audit_id;

            var theURL =  'data/AuditDock.php?command=refresh_audit_status';

            Ext.Ajax.request({

                url: theURL,

                params: {
                    audit_id: this.auditID
                },

                success: function(form, action) {

                    //var theForm = form;
                    //var theAction = action;

                    console.log("configs/AuditStatus.refreshAuditStatus() completed OK");

                },

                failure: function(form, action) {

                    console.log("configs/AuditStatus.refreshAuditStatus() *failed* to complete");

                },

                exception: httpExceptionHandler

            });

        }

        catch(ex) {
            WxLog("error", "configs/AuditStatus.refreshAuditStatus", ex.message);
        }
    },

    saveAuditStatus: function ( edit_status, run_status, status_message ) {

        console.log("configs/AuditStatus.saveAuditStatus(", isDefined(edit_status) ? edit_status : '', ", ",
                                                            isDefined(run_status) ? run_status : '', ", '",
                                                            isDefined(status_message) ? status_message : '', "')");

        if(isDefined(edit_status)){this.editStatus = edit_status;}
        if(isDefined(run_status)){this.runStatus = run_status;}
        if(isDefined(run_status)){this.statusMsg = status_message;}

        var theEntryPoint = 'RetrieveAuditRunStatus';
        var theCommand = 'update';

        var theURL = this.wxWsURL() + '/' + theEntryPoint + '?' +
            'command=' + theCommand + '&' +
            'audit_id=' + this.auditID;

        if(isDefined(edit_status)){
            theURL = theURL + '&' + 'edit_status=' + this.editStatus;
        }

        if(isDefined(run_status)){
            theURL = theURL + '&' + 'run_status=' + this.runStatus;
        }

        if(isDefined(status_message)){
            theURL = theURL + '&' + 'status_message=' + this.statusMsg;
        }

        //WxLog("debug", "configs/AuditStatus.saveAuditStatus()", "URL: "+theURL);

        this.submitRequest( theURL );

    },

    submitRequest: function( theURL ) {

        console.log('configs/AuditStatus.submitRequest()');

        this.requestSubmitted = true;

        //var local_success_callback_function = this.responseReceived;
        var local_failure_callback_function = this.displayErrors;
        var local_http_exception_callback_function = httpExceptionHandler;

        Ext.data.JsonP.request({
            url: theURL,
            method: 'POST',
            callbackKey: "callback",
            disableCaching: true,
            timeout: 300000,
            callback: function(success, result, errorType) {
                wx.configs.AuditStatus.responseReceived(result);
            },
            //success: local_success_callback_function,
            failure: local_failure_callback_function,
            exception: local_http_exception_callback_function
        });

    },

    responseReceived: function(response){

        //console.log('configs/AuditStatus.responseReceived()');

        try {

            var prevEditStatus = this.editStatus;
            var prevRunStatus = this.runStatus;
            var prevStatusMsg = this.statusMsg;

            var theResponse = response;

            if( theResponse.success === "true" ){

                this.editStatus = theResponse.data.edit_status;
                this.runStatus = theResponse.data.run_status;
                this.statusMsg = theResponse.data.status_message;

                if( this.runStatus == 0 && this.editStatus == 2 ){
                    this.runStatus = 1;
                }

                //console.log( 'configs/AuditStatus.responseReceived():  editStatus = ', this.editStatus, ', runStatus = ', this.runStatus, ", statusMsg = '", this.statusMsg, "'" );

            } else {

                console.log( "Run Failed:  Error type = ", theResponse.errors.type, ", error message(s) = ", theResponse.errors.message );

            }

            if( this.requestSubmitted == true || this.editStatus != prevEditStatus ||
                                                 this.runStatus != prevRunStatus ||
                                                 this.statusMsg != prevStatusMsg ) {

                wx.getApplication().fireEvent( 'auditStatusUpdate', theResponse );

            } else {
                console.log('configs/AuditStatus.responseReceived() --- IGNORED!');
            }

            this.requestSubmitted = false;

        }
        catch(ex) {
            WxLog("error", "configs/AuditStatus.responseReceived", ex.message);
        }

    },

    displayErrors: function(response) {

        var msg = 'displaying errors for the request';
        WxLog("debug", "configs/AuditStatus.displayErrors", msg);

        try {

            var is_error = true;
            var middle_div_error_beg = '<div class="h1">Error Report</div><div class"spacer1">&nbsp;</div><ul>';
            var middle_div_result_beg = '<div class="h1">Audit Results</div><div class"spacer1">&nbsp;</div><ul>';
            var middle_div_end = '</ul>';
            var inner_div_beg = '<li type="disc"><div class="error">';
            var inner_div_end = '</div></li>';
            var message_body = inner_div_beg + "An unknown error has occurred." + inner_div_end;
            var msgString = '';

            var theResponse = response;

            if (theResponse === "timeout") {

                message_body = inner_div_beg + "A timeout occurred from the server while processing the request." + inner_div_end;

            } else if (theResponse === "error") {

                message_body = inner_div_beg + "The server has reported an unknown error." + inner_div_end;

            } else if (isDefined(theResponse.errors)) {

                message_body = '<h2>' ;

                if (isDefined(theResponse.errors.auditJobID)) {
                    message_body += '<div class="item"><div class="label">Audit Job: </div><div class="content">' ;
                    message_body += theResponse.errors.auditJobID;
                    message_body += '</div></div><br/>';
                }

                if (isDefined(theResponse.errors.type)) {
                    message_body += '<div class="item"><div class="label">Error Type: </div><div class="content">' ;
                    message_body += theResponse.errors.type;
                    message_body += '</div></div><br/>';
                }

                if (isDefined(theResponse.errors.message)) {
                    message_body += '<div class="error">' ;
                    message_body += theResponse.errors.message;
                    message_body += '</div><br/>';
                }

                message_body += '</h2>';

            } else if (isDefined(theResponse.error)) {
                message_body = '<h2><div class="item"><div class="content">' + theResponse.error + '</div></div></h2><div class"spacer2">&nbsp;</div>';
                message_body += '<div class="item"><div class="content">' + theResponse.message + '</div></div>';
            }

            if (is_error) {
                msgString = middle_div_error_beg + message_body + middle_div_end;
            } else {
                msgString = middle_div_result_beg + message_body + middle_div_end;
            }

            Ext.MessageBox.alert( "An error has occurred ...", msgString );

        }

        catch(ex) {
            WxLog("error", "configs/AuditStatus.displayErrors", ex.message);
        }

        Ext.getBody().unmask();

    }

});