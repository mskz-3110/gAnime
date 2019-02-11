var gAnimeLog = new gEase.Log( "log" );
var gAnimeConfig = {};

function GetAnimePrograms( title, errors ){
  var programs = [];
  title.url = "http://cal.syoboi.jp/tid/"+ title.id +"/time";
  try{
    var options = {
      "method" : "get",
      "muteHttpExceptions" : false,
      "validateHttpsCertificates" : false,
      "followRedirects" : false
    };
    title.request_time = ( new gEase.DateTime() ).ToString();
    var html = UrlFetchApp.fetch( title.url, options ).getContentText( "UTF-8" ).split( "\n" ).join( "" );
    var a = /\<a.*rel="contents".*?\>(.+?)\<\/a\>/m.exec( html );
    do{
      if ( null == a ) break;
      if ( a[ 1 ].indexOf( "アニメ" ) < 0 ){
        break;
      }
      
      var table = /\<table id="ProgList".*?\>(.+?)\<\/table\>/m.exec( html );
      if ( null == table ) break;
      
      var program = null;
      ( new gEase.Regex( /\<td class="(.+?)" *\>(.*?)\<\/td\>/g ) ).Match( table[ 1 ], function( array ){
        switch ( array[ 1 ] ){
          case "ch":{
            var broadcaster = array[ 2 ];
            array = /\>(.+)\</.exec( array[ 2 ] );
            if ( null != array ) broadcaster = array[ 1 ];
            if ( null != program ) programs.push( program );
            
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
            program.start = /\>(.+?)\</.exec( array[ 2 ] )[ 1 ];
          }break;
          case "count":{
            if ( "" != array[ 2 ] ){
              program.count = parseInt( array[ 2 ] );
            }else{
              program.flags.push( "special" );
            }
          }break;
          case "flag":{
            ( new gEase.Regex( /title="(.+?)"/g ) ).Match( array[ 2 ], function( array ){
              switch ( array[ 1 ] ){
                case "新番組":{
                  program.flags.push( "new" );
                }break;
                case "再放送":{
                  program.flags.push( "rerun" );
                }break;
              }
            });
          }break;
        }
      });
      if ( null != program ){
        program.url = title.url;
        program.request_time = title.request_time;
        programs.push( program );
      }
    }while ( false );
  }catch( e ){
    gAnimeLog.E( gEase.json_to_string( [ e.message, title, e.stack ] ) );
    errors.push( title );
  }
  return programs;
}

function GetNewProgramTitles( errors ){
  var titles = [];
  var url = "http://cal.syoboi.jp/quarter/"+ gAnimeConfig.year +"q"+ gAnimeConfig.cours +"?mode=1";
  try{
    var options = {
      "method" : "get",
      "muteHttpExceptions" : false,
      "validateHttpsCertificates" : false,
      "followRedirects" : false
    };
    var html = UrlFetchApp.fetch( url, options ).getContentText( "UTF-8" ).split( "\n" ).join( "" );
    var ol = /\<ol class="titles"\>(.+?)\<\/ol\>/g.exec( html );
    do{
      if ( null == ol ) break;
      
      ( new gEase.Regex( /\<a href="\/tid\/([0-9]+)"\>(.+?)\<\/a\>/g ) ).Match( ol[ 1 ], function( array ){
        titles.push( { "id" : array[ 1 ], "name" : array[ 2 ] } );
      });
    }while ( false );
  }catch( e ){
    gAnimeLog.E( gEase.json_to_string( [ e.message, url, e.stack ] ) );
    errors.push( url );
  }
  return titles;
}

function GetNewAnimePrograms( errors ){
  gEase.each( gEase.sheet( "config" ).getDataRange().getValues(), function( array ){
    var key = array.shift();
    gAnimeConfig[ key ] = array.shift();
  });
  gAnimeLog.D( gEase.json_to_string( gAnimeConfig ) );
  
  var new_anime_programs = gEase.json_from_string( gAnimeConfig[ "new_anime_programs" ], [] );
  do{
    var titles = gEase.json_from_string( gAnimeConfig[ "titles" ], [] );
    if ( ( 0 == titles.length ) && ( 0 == new_anime_programs.length ) ){
      titles = GetNewProgramTitles( errors );
      if ( 0 < errors.length ) break;
      Utilities.sleep( 1000 );
    }
    
    gEase.each( titles, function( title ){
      gEase.each( GetAnimePrograms( title, errors ), function( program ){
        if ( 0 == program.flags.length ) return true;
        
        ( new gEase.Regex( new RegExp( gAnimeConfig.broadcaster_pattern, "g" ) ) ).Match( program.broadcaster, function(){
          var new_anime_program = { "title" : title, "program" : program };
          gAnimeLog.D( gEase.json_to_string( new_anime_program ) );
          new_anime_programs.push( new_anime_program );
        });
      });
      Utilities.sleep( 1000 );
    });
  }while ( false );
  return new_anime_programs;
}

function FlagsToTypes( flags ){
  var types = [];
  gEase.each( flags, function( flag ){
    switch ( flag ){
    case "new":     { types.push( "新番組" ); }break;
    case "special": { types.push( "特別番組" ); }break;
    case "rerun":   { types.push( "再放送" ); }break;
    }
  });
  return types;
}

function Main(){
  var errors = [];
  var new_anime_programs = GetNewAnimePrograms( errors );
  if ( 0 < errors.length ){
    gAnimeLog.E( gEase.json_to_string( errors ) );
  }
  gAnimeLog.I( gEase.json_to_string( new_anime_programs ) );
  
  var sheet = new gEase.Sheet( gEase.sheet( gAnimeConfig.year +"."+ gAnimeConfig.cours ) );
  var row = sheet.GetSheet().getLastRow();
  if ( 0 == row ){
    var record = [ "番組名", "放送局", "開始日時", "種別" ];
    var range = sheet.AddRecord( record );
    range.setFontWeight( "bold" );
    range.setBorder( false, false, true, false, false, false );
    sheet.SetWidths( [ 600, 200, 300, 200 ] );
    ++row;
  }
  gEase.each( new_anime_programs, function( new_anime_program ){
    var record = [ new_anime_program.title.name, new_anime_program.program.broadcaster, new_anime_program.program.start, FlagsToTypes( new_anime_program.program.flags ).join( " " ) ];
    var range = sheet.SetRecord( ++row, 1, record );
    range.setHorizontalAlignment( "left" );
    range.setVerticalAlignment( "middle" );
    range.setFontSize( 16 );
    range.setWrap( false );
    range.setBorder( true, false, true, false, false, false );
    sheet.SetHeight( 50, row );
  });
  sheet.SetFilterAll();
}
