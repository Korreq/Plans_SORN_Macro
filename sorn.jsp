//BEFORE RUNNING MAKE SURE TO CHANGE homeFolder VARIABLE TO LOCATION OF THE CONFIGURATION FILE
var homeFolder = "C:\\Users\\lukas\\Documents\\Github\\Plans_SORN_Macro\\files";

/*  
  Description:
  
  Macro gets nodes' name from input file, then searches through nodes to find connected transformers and generators ( directly or through Y - node ).
  Before any changes original values for each nodes and elements are written with their names into both output files. 
  First file is for voltage change results, second is for reactive power changes.    
  Then for each found generator, increases it's set voltage by value from configuration file and 
  writes it's new value with new calculated values of each element/node into files.       
  Similarly for each transformer except changeing tap by one to increase it's voltage. 
  After succesfully writing results into files, macro shows in a message window time it took to be done.

  Writing nodes names in an input file:
  
  E.g. we have these nodes: ABC111, ABC222, ABC555, ABC444, CAB123, BAC234, BAC235, BAC124, ABCA123, CABC345, ZABC111
  If we want all nodes containing name "ABC" and "BAC" then in input file we can write: 
  
  ABC, bac
  
  White spaces dosen't matter, macro splits input by ',' if we forget to add ',' between searched strings:
  
  ABC bac 
  
  Macro will search for nodes containing "ABCBAC"
  If we want to find nodes ABC111, ABC222 and only them, then we can write:
  
  ABC111, ABC222
  
  Macro will only find ABC111, ABC222 and ZABC111 nodes and not ABC555, ABC444
  In configuration file there are options that can block finding some nodes:
  
  areaId - nodes only from this area will be found, 
  minRatedVoltage - nodes that are rated less than specified will not be found,
  nodeCharIndex and nodeChar - this options are for skiping nodes connecting generators to main nodes e.g. YABC123,
  skipFakeNodes - nodes that don't end on 55 which due to model implementation in plans don't have real representation
*/

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

var node = element = transformer = branch = null;

//Getting variables from config file 
var area = config.areaId, voltage = config.minRatedVoltage, nodeIndex = config.nodeCharIndex, nodeChar = config.nodeChar;

//Setting power flow calculation settings with settings from config file
setPowerFlowSettings( config );

//Calculate power flow, if fails throw error 
CPF();

//Try to read file from location specified in configuration file, then make array from file and close that file
var inputFile = readFile( config, fso );
var inputArray = getInputArray( inputFile );
inputFile.close();

var contains = null;

//Fill node and baseNodesVolt arrays with nodes that matches names from input array
for( var i = 1; i < Data.N_Nod; i++ ){

  contains = false;

  node = NodArray.Get( i );
  
  //If skip fake nodes is set in config, then check if node ends with 55
  if( config.skipFakeNodes && stringContainsAfter( node.Name, "55", strip( node.Name ).length - 2 ) ) continue;
    
  for( var j in inputArray ){

    if( stringContainsWord( strip( node.Name ), inputArray[ j ] ) ){
           
      contains = true;
      
      break;
    }

  }

  //Add node to both arrays that fulfills all conditions:
  //matching area ( if not set to 0 ), connected, not generator's node, higher voltage setpoint than specified in configure file, node contains one of names from input file
  if( ( node.Area === area || node.Area <= 0 ) && node.St > 0 && node.Name.charAt( nodeIndex ) != nodeChar && node.Vn >= voltage && contains ){ 
    
    nodes.push( node );
    
    baseNodesVolt.push( node.Vi );
  }

}

//Fill elements array with valid generators and connected nodes. 
//Also fills baseElementsReactPow with generators reactive power and baseElementsNodesPow with connected nodes power
for( var i = 1; i < Data.N_Gen; i++ ){

  contains = false;

  element = GenArray.Get( i );

  node = NodArray.Get( element.NrNod );
  
  if( element.TrfName ){ 
  
    branch = BraArray.Find( element.TrfName );
    
    node = ( node.Name === branch.EndName ) ? NodArray.Find( branch.BegName ) : NodArray.Find( branch.EndName );
  }
  
  for( var j in inputArray ){

    if( stringContainsWord( strip( node.Name ), inputArray[ j ] ) ){

      contains = true;
      
      break;
    }

  }
 
  //Add valid generators to arrays. Constrains:
  //Minimal reactive power is not equal or higher than maximum reactive power, matches area, 
  //generator's node contains one of names from input file 
  if( element.Qmin < element.Qmax && element.St > 0 && node.Area === area && contains ){
  
    elements.push( [ element, node ] );

    baseElementsReactPow.push( element.Qg );

    baseElementsNodesPow.push( node.Vs );
  }
   
}

//Add valid transformers to arrays with coresponding node and branch. Constrains:
//Transformer must be connected to node from nodes array, have more than 1 tap, not already in elements array
for( i in nodes ){

  node = nodes[ i ];

  for( var j = 1; j < Data.N_Trf; j++ ){

    transformer = TrfArray.Get( j );
    
    if( ( node.Name == transformer.EndName || node.Name == transformer.BegName ) && !elementInArrayByName( elements, transformer.Name ) && transformer.Lstp != 1 ){

      branch = BraArray.Find( transformer.Name );

      elements.push( [ transformer, node, branch ] );
      
      baseElementsNodesPow.push( transformer.Stp0 );
      
      if( node.Name === transformer.EndName ) baseElementsReactPow.push( branch.Qend );
      
      else baseElementsReactPow.push( branch.Qbeg );
    }
    
  }

}

//Create result files and folder with settings from a config file
var file1 = createFile( "Q", config, fso );
var file2 = createFile( "V", config, fso );

//Write headers and base values for each element/node to coresponding file 
file1.Write( "Elements;Old U_G / Tap;New U_G / Tap;" );
file2.Write( "Elements;Old U_G / Tap;New U_G / Tap;" );


//Writes for each element it's node react power 
writeDataToFile( file1, elements, baseElementsReactPow );

//Writes for each element it's node voltage
writeDataToFile( file2, nodes, baseNodesVolt );

//Trying to save file before any change on transformers and connected nodes
if( SaveTempBIN( tmpFile ) < 1 ) errorThrower( "Unable to create temporary file" );

//For each element make some change depending on type of elemenet, then write results into result files
for( i in elements ){ 
  
  var element = elements[ i ][ 0 ], node = elements[ i ][ 1 ];
  
  //If array element have a branch then try to switch tap up 
  if( elements[ i ][ 2 ] ){
  
    if( ( element.TapLoc === 1 && element.Stp0 < element.Lstp ) ) element.Stp0++;
    
    else if( ( element.TapLoc === 0 && element.Stp0 > 1 ) ) element.Stp0--;  
  }

  else{

    //get set value from config file and add it to node's voltage
    var value = config.changeValue;
    
    node.Vs += value;
  }

  //Calculate power flow, if fails try to load original model and throw error 
  if( CalcLF() != 1 ) saveErrorThrower( "Power Flow calculation failed", tmpOgFile );

  //Write element's name, it's base connected node power / tap number and new connected node power / tap number
  if( elements[ i ][ 2 ] ) file1.Write( element.Name + ";" + baseElementsNodesPow[ i ] + ";" + element.Stp0 + ";" );
  
  else file1.Write( element.Name + ";" + roundTo( baseElementsNodesPow[ i ], 2 ) + ";" + roundTo( node.Vs, 2 ) + ";" );

  var react = null; 
  
  //Write for each element it's new reactive power
  for( j in elements ){

    react = null;

    //Check if element have a branch, if true use reactive power from matching branch end
    if( elements[ j ][ 2 ] ){

      react = ( elements[ j ][ 0 ].begName === elements[ j ][ 1 ].Name ) ? elements[ j ][ 2 ].Qbeg : elements[ j ][ 2 ].Qend;
    } 

    else react = elements[ j ][ 0 ].Qg;
  
    file1.Write( roundTo( react, 2 ) + ";" );
  }
  
  //Add end line character to file
  file1.WriteLine("");

  //Write element's name, it's base connected node power / tap number and new connected node power / tap number
  if( elements[ i ][ 2 ] ) file2.Write( element.Name + ";" + baseElementsNodesPow[ i ] + ";" + element.Stp0 + ";" );
  
  else file2.Write( element.Name + ";" + roundTo( baseElementsNodesPow[ i ], 2 ) + ";" + roundTo( node.Vs, 2 ) + ";" );
  
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

//Gets program working duration and shows it in a message box
//Working duration dosen't work if program starts and ends in a different day due to lack of futher date checking 
time = getTime() - time;
MsgBox( "Task completed in " + formatTime( time ), 0 | 64, "Task completed" );

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

//Function takes string and returns it without whitespaces
function strip( string ){

  var strippedString = string;
  
  return strippedString.replace(/(^\s+|\s+$)/g, '');
}

//Function checks if searched word is in a string, can change from where to start checking for match
function stringContainsAfter( string, word, start ){

  var j = 0;

  for( var i = start; i < string.length; i++ ){
  
    j = ( string.charAt( i ) === word.charAt( j ) ) ? j + 1 : 0;

    if( j === word.length ) return true; 
  }
  
  return false;
}

//Function checks if searched word is in a string
function stringContainsWord( string, word ){
  
  return stringContainsAfter( string, word, 0 );
}

//Function gets each element's name from 2D array and compares it to elementName 
function elementInArrayByName( array, elementName ){

  for( i in array ){
  
    if(array[ i ][ 0 ].Name === elementName) return true;
  }

  return false;
}

//Function writes to specifed file objects names and corresponding data
function writeDataToFile( file, objectArray, dataArray ){

  var text = "Base;X;X;";
  
  for( i in objectArray ){
  
    if( objectArray[ i ][ 0 ] ) file.Write( objectArray[ i ][ 0 ].Name + ";" );

    else file.Write( objectArray[ i ].Name + ";" );
    
    text += roundTo( dataArray[ i ], 2 ) + ";";
  }

  file.WriteLine( "\n" + text );
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
    skipFakeNodes: ini.GetBool( "variable", "skipFakeNodes", 0 ),

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
  ini.WriteBool( "variable", "skipFakeNodes", conf.skipFakeNodes );

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

//Function gets file and takes each line into a array and after finding "," character pushes array into other array. Returns string array
function getInputArray( file ){

  var array = []; 

  while(!file.AtEndOfStream){

    var tmp = [], line = file.ReadLine(), word = null;
    
    tmp = line.split(",");

    for( i in tmp ){
      
      word = tmp[ i ].replace(/(^\s+|\s+$)/g, '');
      
      if( word != "" ) array.push( word.toUpperCase() );
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

//Function returns current time in seconds 
function getTime(){
  
  var current = new Date();
  
  return current.getHours() * 3600 + current.getMinutes() * 60 + current.getSeconds();
}

//Function takes time in seconds and returns time in HH:MM:SS format
function formatTime( time ){

  var hours = minutes = 0;

  hours = Math.floor( time / 3600 );

  time -= hours * 3600;

  minutes = Math.floor( time / 60 );

  time -= minutes * 60;

  return ( '0' + hours ).slice( -2 ) + ":" + ( '0' + minutes ).slice( -2 ) + ":" + ( '0' + time ).slice( -2 );
}
