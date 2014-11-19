<?php

require_once "dataHelper.php";

function updateUser( $sprocName, $data, $dataName ) {

    $args = implode( ", ", array_values($data));
    
    $sqlString = "CALL `".$sprocName."`( ".$args." );";

    $results = CallGenericSelectIQuery($sqlString, $dataName);
    
    return $results;
}

try {

    CheckSession();

    $command = strtolower($_GET["command"]);

    $editingUserId = getParam('editing_user_id');

    $user_id = getParam('user__user_id');
    $contact_id = getParam('user__contact_id');
    $agency_id = getParam('user__agency_id');

    $dataName = "data";
    $fieldsNamePrefix = "user";
    $fieldsDelim = "__";

    switch($command) {

        case "list":

            $sprocName = "user_list_narrowed_select";

            $sqlString = "CALL `".$sprocName."`( ".$editingUserId." )";

            $results = CallGenericSelectQuery( $sqlString, $dataName );

            echo json_encode( PrefaceArrayDataKeys( $results, $fieldsNamePrefix, $fieldsDelim, $dataName ) );

            break;


        case "select":

            $sprocName = "user_select";

            $results = CallGenericJobSelectProcedure( $sprocName, $user_id, $dataName );

            echo json_encode( PrefaceArrayDataKeys( $results, $fieldsNamePrefix, $fieldsDelim, $dataName ) );

            break;


        case "create":
        case "update":

            $sprocName = 'user_update';

            $fieldsQuote = array( 'user_name', 'user_pword', 'user_comments',
                                  'first_name', 'middle_initial', 'last_name', 'company', 'title',
                                  'street_address', 'unit_number', 'city', 'zip_code',
                                  'work_phone', 'cell_phone', 'home_phone', 'fax_phone', 'email_address', 'web_address', 'contact_comments' );

            $fieldsNullIfBlank = array( );

            $immutableFields = array( );

            $sprocArgs = GetPostedSprocData( $sprocName, $fieldsNamePrefix, $fieldsDelim, $fieldsQuote, $fieldsNullIfBlank, $immutableFields, $dataName );

            $results = updateUser( $sprocName, $sprocArgs, $dataName);

            echo json_encode( ProcessResultClose( $results ) );

            break;


        case "delete":

            $sprocName = "user_delete";

            $sqlString = "CALL `".$sprocName."`( ".$user_id.", ".$agency_id." );";

            $results = CallGenericSelectIQuery( $sqlString, $dataName );


            break;


        case "getusernamecount":

            $user_name = getParam('user__user_name');

            $sqlString = "SELECT COUNT(*) AS 'num_matching' FROM `user` WHERE `user_name` = '".$user_name."';";

            $results = CallGenericSelectQuery( $sqlString, $dataName );

            echo json_encode($results);

            break;


        default:

            $errorMsg = "unrecognized command: ".$command;

            LogString( "User.php - unknown command: ".$command, $errorMsg );

            die( array('success'=>false,'msg'=>$errorMsg));

            break;

    }

}

catch( Exception $e) {

    LogString( "User.php - Exception: ".$e->getCode(), $e->getMessage() );

    die( array( 'success'=>false, 'msg'=>$e->getMessage() ) );

}

?>