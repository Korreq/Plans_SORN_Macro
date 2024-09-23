//
//TODO skip nodes ending on 55 eg. GBL455, not real nodes
//

//Location of folder where config file is located
var homeFolder = "C:\\Users\\lukas\\Documents\\Github\\Plans_GeneratorsVoltageChangeTest_Macro\\files";

//Creating file operation object
var fso = new ActiveXObject( "Scripting.FileSystemObject" );

//Initializing configuration object
var config = iniConfigConstructor( homeFolder, fso );
var tmpFile = config.homeFolder + "\\tmp.bin", tmpOgFile = config.homeFolder + "\\tmpOrg.bin";

//Loading kdm model file and trying to save it as temporary binary file
ReadDataKDM( config.modelPath + "\\" + config.modelName + ".kdm" );
if( SaveTempBIN( tmpOgFile ) < 1 ) errorThrower( "Unable to create temporary file" );

var time = getTime();

var nodes = [], elements = [];

var baseElementsReactPow = [], baseNodesVolt = [], baseElementsNodesPow = [];

//Getting variables from config file 
var area = config.areaId, voltage = config.minRatedVoltage, nodeIndex = config.nodeCharIndex, nodeChar = config.nodeChar;

//Setting power flow calculation settings with settings from config file
setPowerFlowSettings( config );

//Calculate power flow, if fails throw error 
CPF();

//Try to read file from location specified in configuration file, then make array from file and close it
var inputFile = readFile( config, fso );
var inputArray = getInputArray( inputFile );
inputFile.close();

var contains = null;

//Fill node and baseNodesVolt arrays with nodes that matches names from input file
for( var i = 1; i < Data.N_Nod; i++ ){

  contains = false;

  var n = NodArray.Get( i );

  //Add to config file
  if( stringContainsWord( n.Name, "55" ) ) continue;
  
  for( var j in inputArray ){

    if( stringContainsWord( n.Name, inputArray[ j ] ) ){
      
      contains = true;
      
      break;
    }

  }

  //Add node to both arrays that fulfills all conditions:
  //matching area, connected, not generator's node, higher voltage setpoint than specified in configure file, node contains one of names from input file
  if( n.Area === area && n.St > 0 && n.Name.charAt( nodeIndex ) != nodeChar && n.Vn >= voltage && contains ){ 
    
    nodes.push( n );
    
    baseNodesVolt.push( n.Vi );
  }

}

//Fill elements array with valid generators and connected nodes. 
//Also fills baseElementsReactPow with generators reactive power and baseElementsNodesPow with connected nodes power
for( var i = 1; i < Data.N_Gen; i++ ){

  contains = false;

  var g = GenArray.Get( i );

  var n = NodArray.Get( g.NrNod );

  for( var j in inputArray ){

    if( stringContainsWord( n.Name, inputArray[ j ] ) ){
      
      contains = true;
      
      break;
    }

  }
 
  //Add valid generators to arrays. Constrains:
  //Minimal reactive power is not equal or higher than maximum reactive power, generator is connected to grid, matches area, 
  //generator's node contains one of names from input file 
  if( g.Qmin < g.Qmax && g.St > 0 && n.Area === area && n.Name.charAt( nodeIndex ) == nodeChar && contains ){

    elements.push( [ g, n ] );

    baseElementsReactPow.push( g.Qg );

    baseElementsNodesPow.push( n.Vs );
  }
   
}

var b = null;

//Add valid transformers to arrays with coresponding node and branch. Constrains:
//Transformer must be connected to node from nodes array, have more than 1 tap, not already been in elements array
for( i in nodes ){

  var n = nodes[ i ];

  for( var j = 1; j < Data.N_Trf; j++ ){

    var t = TrfArray.Get( j );
    
    if( ( n.Name == t.EndName || n.Name == t.BegName ) && !elementInArrayByName( elements, t.Name ) && t.Lstp != 1 ){

      b = BraArray.Find( t.Name );

      elements.push( [ t, n, b ] );
      
      baseElementsNodesPow.push( t.Stp0 );
      
      if( n.Name === t.EndName ) baseElementsReactPow.push( b.Qend );
      
      else baseElementsReactPow.push( b.Qbeg );
      
    }
    
  }

}

//Create result files and folder with settings from a config file
var file1 = createFile( "Q", config, fso );
var file2 = createFile( "V", config, fso );

//Write headers and base values for each element/node to coresponding file 
file1.Write( "Elements;Old U_G / Tap;New U_G / Tap;" );
file2.Write( "Elements;Old U_G / Tap;New U_G / Tap;" );

var temp = "Base;X;X;";

for( i in elements ){

  file1.Write( elements[ i ][ 0 ].Name + ";" );

  temp += roundTo( baseElementsReactPow[ i ], 2 ) + ";";
}

file1.WriteLine( "\n" + temp );

temp = "Base;X;X;";

for( i in nodes ){

  file2.Write( nodes[ i ].Name + ";" );

  temp += roundTo( baseNodesVolt[ i ], 2 ) + ";";
}

file2.WriteLine( "\n" + temp );

//Trying to save file before any change on transformers and connected nodes
if( SaveTempBIN( tmpFile ) < 1 ) errorThrower( "Unable to create temporary file" );

for( i in elements ){
  
  var g = elements[ i ][ 0 ], n = elements[ i ][ 1 ];
  
  //Check if generator has block transformer
  if( g.TrfName != "" && !elements[ i ][ 2 ] ){

    //Find transformer and change it's type to 11 ( without regulation )
    var t = TrfArray.Find( g.TrfName );
    t.Typ = 11;
    
    //Check if transformer name's ends with A, indicates that there are more than 1 block transformers connected to generator
    if( g.TrfName.charAt( g.TrfName.length - 1 ) == 'A' ){
      
      var l = 'B';
      //Transformer name without last char
      var tName = g.TrfName.slice(0, -1);
      
      //As long as there are transformers with same name ending with next letter, then change it's type to 11 ( without regulation )
      while( true ){
        
        //Try to assign a transformer to variable and transformer type to 11 ( without regulation )
        try{
          
          t = TrfArray.Find( tName + l ); 
          t.Typ = 11;
        }
        
        //If t is null then exit while loop
        catch( e ){ break; }
        
        //changes l to next letter by adding 1 to it's char code
        l = String.fromCharCode ( l.charCodeAt( 0 ) + 1 );
      }

    }

  }

  //If array element have a branch then try to switch tap up 
  if( elements[ i ][ 2 ] ){
  
    if( ( g.TapLoc === 1 && g.Stp0 < g.Lstp ) ) g.Stp0++;
    
    else if( ( g.TapLoc === 0 && g.Stp0 > 1 ) ) g.Stp0--;  
  }

  else{

    //get set value from config file and add it to node's voltage
    var value = config.changeValue;
    
    n.Vs += value;
  }

  //Calculate power flow, if fails try to load original model and throw error 
  if( CalcLF() != 1 ) saveErrorThrower( "Power Flow calculation failed", tmpOgFile );

  //Write element's name, it's base connected node power / tap number and new connected node power / tap number
  if( elements[ i ][ 2 ] ) file1.Write( g.Name + ";" + baseElementsNodesPow[ i ] + ";" + g.Stp0 + ";" );
  
  else file1.Write( g.Name + ";" + roundTo( baseElementsNodesPow[ i ], 2 ) + ";" + roundTo( n.Vs, 2 ) + ";" );

  var react = null; 
  
  //Write for each element it's new reactive power
  for( j in elements ){

    react = null;

    if( elements[ j ][ 2 ] ){

      react = ( elements[ j ][ 0 ].begName === elements[ j ][ 1 ].Name ) ? elements[ j ][ 2 ].Qbeg : elements[ j ][ 2 ].Qend;
    } 

    else react = elements[ j ][ 0 ].Qg;
  
    file1.Write( roundTo( react, 2 ) + ";" );
  }
  
  //Add end line character to file
  file1.WriteLine("");

  //Write element's name, it's base connected node power / tap number and new connected node power / tap number
  if( elements[ i ][ 2 ] ) file2.Write( g.Name + ";" + baseElementsNodesPow[ i ] + ";" + g.Stp0 + ";" );
  
  else file2.Write( g.Name + ";" + roundTo( baseElementsNodesPow[ i ], 2 ) + ";" + roundTo( n.Vs, 2 ) + ";" );
  
  //Write for each node it's new voltage
  for( j in nodes ) file2.Write( roundTo( nodes[ j ].Vi, 2 ) + ";" );

  //Add end line character to file
  file2.WriteLine("");

  //Load model without any changes to transformators
  ReadTempBIN( tmpFile );
}

//Loading original model
ReadTempBIN( tmpOgFile );

//Removing temporary binary files
fso.DeleteFile( tmpFile );
fso.DeleteFile( tmpOgFile );

//Closing result files
file1.Close();
file2.Close();

time = getTime() - time;
cprintf( "Time Elapsed: "+ time ); 

//Function uses JS Math.round, takes value and returns rounded value to specified decimals 
function roundTo( value, precision ){

  return Math.round( value * ( 10 * precision ) ) / ( 10 * precision ) ;
}

//Set power flow settings using config file
function setPowerFlowSettings( config ){

  Calc.Itmax = config.maxIterations;
  Calc.EPS10 = config.startingPrecision;
  Calc.Eps = config.precision;
  Calc.EpsUg = config.uzgIterationPrecision;
  Calc.Met = config.method;
}

//Function try to load original file before throwing an error
function saveErrorThrower( message, binPath ){

  try{ ReadTempBIN( binPath ); }
  
  catch( e ){ MsgBox( "Couldn't load original model", 16 ) }

  errorThrower( message );
}

//Basic error thrower with error message window
function errorThrower( message ){
  
  MsgBox( message, 16, "Error" );

  throw message;
}

//Calls built in power flow calculate function, throws error if it fails
function CPF(){

  if( CalcLF() != 1 ) errorThrower( "Power Flow calculation failed" );
}

//Function checks if word is in a string. Word can only be matched whole
function stringContainsWord( string, word ){
  
  var j = 0;

  for( var i = 0; i < string.length; i++ ){
  
    j = ( string.charAt( i ) === word.charAt( j ) ) ? j + 1 : 0;
  
    if( j === word.length ) return true;
  }
  
  return false;
}

//Function gets each element's name from 2D array and compares it to elementName 
function elementInArrayByName( array, elementName ){

  for( i in array ){
  
    if(array[ i ][ 0 ].Name === elementName) return true;
  }

  return false;
}

//Function takes config object and depending on it's config creates folder in specified location. 
//Throws error if config object is null and when folder can't be created
function createFolder( config, fso ){
  
  if( !config ) errorThrower( "Unable to load configuration" );
  
  var folder = config.folderName;
  var folderPath = config.homeFolder + folder;
  
  if( !fso.FolderExists( folderPath ) ){
    
    try{ fso.CreateFolder( folderPath ); }
    
    catch( err ){ errorThrower( "Unable to create folder" ); }
  }
  
  folder += "\\";

  return folder;
}

//Function takes config object and depending on it's config creates file in specified location.
//Also can create folder where results are located depending on configuration file 
//Throws error if config object is null and when file can't be created
function createFile( name, config, fso ){
 
  if( !config ) errorThrower( "Unable to load configuration" );

  var file = null;
  
  var folder = ( config.createResultsFolder == 1 ) ? createFolder( config, fso ) : "";
  var timeStamp = ( config.addTimestampToResultsFiles == 1 ) ? getCurrentDate() + "--" : "";
  var fileLocation = config.homeFolder + folder + timeStamp + config.resultsFilesName + "--" + name + ".csv";
  
  try{ file = fso.CreateTextFile( fileLocation ); }
  
  catch( err ){ errorThrower( "File already exists or unable to create it" ); }

  return file;
} 

//Function takes config object and depending on it reads file from specified location.
//Throws error if config object is null or when file can't be read 
function readFile( config, fso ){

  if( !config ) errorThrower( "Unable to load configuration" );

  var file = null;

  var fileLocation = config.inputFileLocation + config.inputFileName + "." + config.inputFileFormat;
  
  try{ file = fso.OpenTextFile( fileLocation, 1, false, 0 ); }

  catch( err ){ errorThrower( "Unable to find or open file" ); }

  return file;
}

//Function uses built in .ini function to get it's settings from config file.
//Returns conf object with settings taken from file. If file isn't found error is throwed instead.
function iniConfigConstructor( iniPath, fso ){
  
  var configFile = iniPath + "\\config.ini";

  if( !fso.FileExists( configFile ) ) errorThrower( "config.ini file not found" );

  //Initializing plans built in ini manager
  var ini = CreateIniObject();
  ini.Open( configFile );

  var hFolder = ini.GetString( "main", "homeFolder", Main.WorkDir );
  
  //Declaring conf object and trying to fill it with config.ini configuration
  var conf = {
  
    //Main
    homeFolder: hFolder,
    modelName: ini.GetString( "main", "modelName", "model" ),
    modelPath: ini.GetString( "main", "modelPath", hFolder ),  
   
    //Variable
    areaId: ini.GetInt( "variable", "areaId", 1 ),
    minRatedVoltage: ini.GetInt( "variable", "minRatedVoltage", 0 ),
    nodeCharIndex: ini.GetInt( "variable", "nodeCharIndex", 0 ),
    nodeChar: ini.GetString( "variable", "nodeChar", 'Y' ),
    changeValue: ini.GetInt( "variable", "changeValue", 1 ),

    //Folder
    createResultsFolder: ini.GetBool( "folder", "createResultsFolder", 0 ),
    folderName: ini.GetString( "folder", "folderName", "folder" ),
    
    //Files
    inputFileLocation: ini.GetString( "files", "inputFileLocation", hFolder ),
    inputFileName: ini.GetString( "files", "inputFileName", "input" ),
    inputFileFormat: ini.GetString( "files", "inputFileFormat", "txt" ),
    addTimestampToResultsFiles: ini.GetBool( "files", "addTimestampToResultsFiles", 1 ),
    resultsFilesName: ini.GetString( "files", "resultsFilesName", "result" ),
    roundingPrecision: ini.GetInt( "files", "roundingPrecision", 2 ),
    
    //Power Flow
    maxIterations: ini.GetInt( "power flow", "maxIterations", 300 ),
    startingPrecision: ini.GetDouble( "power flow", "startingPrecision", 10.00 ),
    precision: ini.GetDouble( "power flow", "precision", 1.00 ),
    uzgIterationPrecision: ini.GetDouble( "power flow", "uzgIterationPrecision", 0.001 ),
    method: ini.GetInt( "power flow", "method", 1 )
  };
  
  //Overwriting config.ini file
  //Main
  ini.WriteString( "main", "homeFolder", conf.homeFolder );
  ini.WriteString( "main", "modelName", conf.modelName );
  ini.WriteString( "main", "modelPath", conf.modelPath );
  
  //Variable
  ini.WriteInt( "variable", "areaId", conf.areaId );
  ini.WriteInt( "variable", "minRatedVoltage", conf.minRatedVoltage );
  ini.WriteInt( "variable", "nodeCharIndex", conf.nodeCharIndex );
  ini.WriteString( "variable", "nodeChar", conf.nodeChar );
  ini.WriteInt( "variable", "changeValue", conf.changeValue );

  //Folder
  ini.WriteBool( "folder", "createResultsFolder", conf.createResultsFolder );
  ini.WriteString( "folder", "folderName", conf.folderName );
    
  //Files
  ini.WriteString( "files", "inputFileLocation", conf.inputFileLocation );
  ini.WriteString( "files", "inputFileName", conf.inputFileName );
  ini.WriteString( "files", "inputFileFormat", conf.inputFileFormat );
  ini.WriteBool( "files", "addTimestampToResultsFiles", conf.addTimestampToResultsFiles );
  ini.WriteString( "files", "resultsFilesName", conf.resultsFilesName );
  ini.WriteInt( "files", "roundingPrecision", conf.roundingPrecision );
    
  //Power Flow
  ini.WriteInt( "power flow", "maxIterations", conf.maxIterations );
  ini.WriteDouble( "power flow", "startingPrecision", conf.startingPrecision );
  ini.WriteDouble( "power flow", "precision", conf.precision );
  ini.WriteDouble( "power flow", "uzgIterationPrecision", conf.uzgIterationPrecision );
  ini.WriteInt( "power flow", "method", conf.method );
 
  return conf;
}

//Function gets file and takes each line into a array and after finding "," character pushes array into other array
function getInputArray( file ){

  var array = []; 

  while(!file.AtEndOfStream){

    var tmp = [], line = file.ReadLine(), word = null;
    
    tmp = line.split(",");

    for( i in tmp ){
      
      word = tmp[ i ].replace(/(^\s+|\s+$)/g, '');
      
      if( word != "" ) array.push( word );
    }

  }

  return array;
}

//Function takes current date and returns it in file safe format  
function getCurrentDate(){
  
  var current = new Date();
  
  var formatedDateArray = [ ( '0' + ( current.getMonth() + 1 ) ).slice( -2 ), ( '0' + current.getDate() ).slice( -2 ), 
  ( '0' + current.getHours() ).slice( -2 ), ( '0' + current.getMinutes() ).slice( -2 ), ( '0' + current.getSeconds() ).slice( -2 ) ];
  
  return current.getFullYear() + "-" + formatedDateArray[ 0 ] + "-" + formatedDateArray[ 1 ] + "--" + formatedDateArray[ 2 ] + "-" + formatedDateArray[ 3 ] + "-" + formatedDateArray[ 4 ];
}

function getTime(){
  
  var current = new Date();
  
  return current.getHours() * 3600 + current.getMinutes() * 60 + current.getSeconds();
}