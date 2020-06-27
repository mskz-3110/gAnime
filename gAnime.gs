var gAnimeLog = new gEase.Log( "log" );
var gAnimeConfig = {};

function GetAnimePrograms( title ){
  var result = {
    "url" : "http://cal.syoboi.jp/tid/"+ title.id +"/time",
    "request_time" : ( new gEase.DateTime() ).ToString(),
    "programs" : [],
    "errors" : []
  };
  try{
    var options = {
      "method" : "get",
      "muteHttpExceptions" : false,
      "validateHttpsCertificates" : false,
      "followRedirects" : true
    };
    var html = UrlFetchApp.fetch( result.url, options ).getContentText( "UTF-8" ).split( "\n" ).join( "" );
    var a = /\<a.*rel="contents".*?\>(.+?)\<\/a\>/m.exec( html );
    do{
      if ( null == a ) break;
      if ( a[ 1 ].indexOf( "アニメ" ) < 0 ) break;
      
      var table = /\<table id="ProgList".*?\>(.+?)\<\/table\>/m.exec( html );
      if ( null == table ) break;
      
      var program = null;
      ( new gEase.Regex( /\<td class="(.+?)" *\>(.*?)\<\/td\>/g ) ).Match( table[ 1 ], array => {
        switch ( array[ 1 ] ){
        case "ch":{
          var broadcaster = array[ 2 ];
          array = /\>(.+)\</.exec( array[ 2 ] );
          if ( null != array ) broadcaster = array[ 1 ];
          if ( null != program ) result.programs.push( program );
          
          program = {
            "broadcaster" : broadcaster,
            "minites" : 0,
            "start" : "",
            "count" : 0,
            "flags" : []
          };
        }break;
        
        case "min":{
          program.minites = parseInt( array[ 2 ] );
        }break;
        
        case "start":{
          program.start = /\>(.+?)\</.exec( array[ 2 ] )[ 1 ] +" ";
          var start_time = /; *([0-9\:]+)/.exec( array[ 2 ] );
          if ( start_time[ 1 ].length < 5 ) start_time[ 1 ] = "0" + start_time[ 1 ];
          program.start += start_time[ 1 ];
        }break;
        
        case "count":{
          if ( "" != array[ 2 ] ){
            program.count = parseInt( array[ 2 ] );
          }else{
            program.flags.push( "special" );
          }
        }break;
        
        case "flag":{
          ( new gEase.Regex( /title="(.+?)"/g ) ).Match( array[ 2 ], array => {
            switch ( array[ 1 ] ){
            case "新番組":{ program.flags.push( "new" ); }break;
            case "再放送":{ program.flags.push( "rerun" ); }break;
            }
          });
        }break;
        }
      });
      if ( null != program ) result.programs.push( program );
    }while ( false );
  }catch( e ){
    gAnimeLog.E( gEase.json_to_string( [ e.message, title, result.url, result.reqeust_time, e.stack ] ) );
    result.errors.push( title );
  }
  return result;
}

function GetNewAnimeProgramTitles( year, cours ){
  var result = {
    "url" : "http://cal.syoboi.jp/quarter/"+ year +"q"+ cours +"?mode=1",
    "titles" : [],
    "errors" : []
  };
  try{
    var options = {
      "method" : "get",
      "muteHttpExceptions" : false,
      "validateHttpsCertificates" : false,
      "followRedirects" : true
    };
    var html = UrlFetchApp.fetch( result.url, options ).getContentText( "UTF-8" ).split( "\n" ).join( "" );
    var ol = /\<ol class="titles"\>(.+?)\<\/ol\>/g.exec( html );
    do{
      if ( null == ol ) break;
      
      ( new gEase.Regex( /\<a href="\/tid\/([0-9]+)"\>(.+?)\<\/a\>/g ) ).Match( ol[ 1 ], array => {
        result.titles.push( { "id" : array[ 1 ], "name" : array[ 2 ] } );
      });
    }while ( false );
  }catch( e ){
    gAnimeLog.E( gEase.json_to_string( [ e.message, result.url, e.stack ] ) );
    result.errors.push( url );
  }
  return result;
}

function GetNewAnimePrograms(){
  var result = {
    "programs" : [],
    "errors" : []
  };
  gEase.sheet_get( "config" ).getDataRange().getValues().forEach( array => {
    var key = array.shift();
    gAnimeConfig[ key ] = array.shift();
  });
  gAnimeLog.D( gEase.json_to_string( gAnimeConfig ) );
  
  do{
    var titles = gEase.json_from_string( gAnimeConfig[ "titles" ], [] );
    if ( 0 == titles.length ){
      var new_anime_program_titles = GetNewAnimeProgramTitles( gAnimeConfig.year, gAnimeConfig.cours );
      if ( 0 < new_anime_program_titles.length ){
        result.errors = new_anime_program_titles.errors;
        break;
      }
      titles = new_anime_program_titles.titles;
      Utilities.sleep( 1000 );
    }
    
    titles.forEach( title => {
      if ( title.title ) title = title.title;
      var anime_programs = GetAnimePrograms( title );
      if ( 0 < anime_programs.errors.length ) result.errors = result.errors.concat( anime_programs.errors );
      anime_programs.programs.forEach( program => {
        if ( 0 < program.flags.length ){
          ( new gEase.Regex( new RegExp( gAnimeConfig.broadcaster_pattern, "g" ) ) ).Match( program.broadcaster, _ => {
            var new_anime_program = { "title" : title, "program" : program };
            gAnimeLog.D( gEase.json_to_string( new_anime_program ) );
            result.programs.push( new_anime_program );
          });
        }
      });
      Utilities.sleep( 1000 );
    });
  }while ( false );
  return result;
}

function FlagsToTypes( flags ){
  var types = [];
  flags.forEach( flag => {
    switch ( flag ){
    case "new":     { types.push( "新番組" ); }break;
    case "special": { types.push( "特別番組" ); }break;
    case "rerun":   { types.push( "再放送" ); }break;
    }
  });
  return types;
}

function Main(){
  var new_anime_programs = GetNewAnimePrograms();
  if ( 0 < new_anime_programs.errors.length ){
    gAnimeLog.E( gEase.json_to_string( new_anime_programs.errors ) );
  }
  gAnimeLog.I( gEase.json_to_string( new_anime_programs.programs ) );
  
  var sheet = new gEase.Sheet( gEase.sheet_get_or_add( gAnimeConfig.year +"."+ gAnimeConfig.cours ) );
  var row = sheet.GetSheet().getLastRow();
  if ( 0 == row ){
    var record = [ "番組名", "放送局", "開始日時", "種別" ];
    var range = sheet.AddRecord( 1, record );
    range.setFontWeight( "bold" );
    range.setBorder( false, false, true, false, false, false );
    sheet.SetWidths( 1, [ 600, 200, 300, 200 ] );
    ++row;
  }
  new_anime_programs.programs.forEach( new_anime_program => {
    var record = [ gEase.html_decode( new_anime_program.title.name ), new_anime_program.program.broadcaster, new_anime_program.program.start, FlagsToTypes( new_anime_program.program.flags ).join( " " ) ];
    var range = sheet.SetRecord( ++row, 1, record );
    range.setHorizontalAlignment( "left" );
    range.setVerticalAlignment( "middle" );
    range.setFontSize( 16 );
    range.setWrap( false );
    range.setBorder( true, false, true, false, false, false );
    sheet.SetHeight( row, 50 );
  });
  sheet.SetFilterAll();
}
