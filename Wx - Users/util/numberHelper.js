function isNumeric(value) {

    try {
        return ( !isNaN(parseFloat(value)) && isFinite(value) );
    }

    catch(ex) {
        console.log('a call to isNumeric(', value, ') encountered an error:');
        console.log(ex);
        return false;
    }

}

var ones=['','one','two','three','four','five','six','seven','eight','nine'];
var tens=['','','twenty','thirty','forty','fifty','sixty','seventy','eighty','ninety'];
var teens=['ten','eleven','twelve','thirteen','fourteen','fifteen','sixteen','seventeen','eighteen','nineteen'];

function convert_millions( num ){

    var numAsWords = '';

    if ( num >= 1000000 ) {
        numAsWords = convert_millions( Math.floor( num / 1000000 ) ) + " million " + convert_thousands( num % 1000000 );
    } else {
        numAsWords = convert_thousands( num );
    }

    return numAsWords;

}

function convert_thousands( num ){

    var numAsWords = '';

    if ( num >= 1000 ) {
        numAsWords = convert_hundreds( Math.floor( num/1000 ) ) + " thousand " + convert_hundreds( num % 1000 );
    } else {
        numAsWords = convert_hundreds( num );
    }

    return numAsWords.trim();

}

function convert_hundreds( num ){

    var numAsWords = '';

    if ( num > 99 ) {
        numAsWords = ones[ Math.floor( num / 100 ) ] + " hundred " + convert_tens( num % 100 );
    } else {
        numAsWords = convert_tens( num );
    }

    return numAsWords.trim();

}

function convert_tens( num ){

    var numAsWords = '';

    if ( num < 10 ) {
        numAsWords = ones[ num ];
    } else  if ( num >= 10 && num < 20 ) {
        numAsWords = teens[ num - 10 ];
    } else {
        numAsWords = tens[ Math.floor( num / 10 ) ] + " " +ones[ num % 10 ];
    }

    return numAsWords.trim();

}

function toWords( num ){

    var numAsWords = 'Not a number';
    var valuePrefix = '';

    if ( isNumeric( num ) ) {

        if ( num < 0 ) {
            valuePrefix = 'negative ';
            num = Math.abs( num );
        }

        if ( num == 0 ) {
            numAsWords = "zero";
        } else {
            numAsWords = convert_millions( num );
        }

    }

    return valuePrefix + numAsWords;

}



//testing code begins here

function toWordsTest(){
    var cases=[-123456,-789,-1,0,1,2,7,10,11,12,13,15,19,20,21,25,29,30,35,50,55,69,70,99,100,101,119,510,900,1000,5001,5019,5555,10000,11000,100000,199001,1000000,1111111,190000009];
    for (var i=0;i<cases.length;i++ ){
        console.log( "toWords(", cases[i], ") = ", toWords(cases[i]) );
    }
}
