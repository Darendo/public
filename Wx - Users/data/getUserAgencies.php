<?php

include_once "dataHelper.php";

function GetSprocResults( $queryString ) {

    $enum_values = array( 'data' => array(), 'success' => true ); //Adding in success for JSON

    $conn = OpenDbIConnection();

    $qryResults = mysqli_query($conn, $queryString, MYSQLI_USE_RESULT);

    $i = 0;

    while ($row = mysqli_fetch_assoc($qryResults)) {

        $enum_values['data'][$i] = $row;
        $i++;

    }

    CloseDbIConnection($conn);

    return $enum_values;

}

try {

    CheckSession();

    $sproc_name = getParam( 'sproc', '' );
    $user_id = getParam( 'user_id', 0 );

    if( $sproc_name !== '' ) {

        if( $user_id > 0 ) {

            $qryString = "CALL `".$sproc_name."`( ".$user_id." );";

            $results = GetSprocResults($qryString);

        } else {

            $results = array( 'data' => array(), 'success' => false );

        }

    } else {

        $results = array( 'data' => array(), 'success' => false );

    }

    echo json_encode($results);

}

catch( Exception $e) {

    LogString( "Exception=".$e->getCode(), $e->getMessage() );

    die( array( 'success'=>false, 'msg'=>$e->getMessage() ) );

}