// BEFORE RUNNING MAKE SURE TO CHANGE configurationFileLocation VARIABLE TO LOCATION OF THE CONFIGURATION FILE
var configurationFileLocation = "Z:\\home\\yoga\\Documents\\Github\\Plans_SORN_Macro\\files";

/* 
  Macro Overview:
  This script retrieves node names from an input file, searches for connected transformers and generators, 
  and writes original values to output files. It then updates generator voltages and transformer taps, 
  recalculates values, and writes the results to output files.


  Input File Format:
  Node names are searched using a case-insensitive regex pattern. 
  Multiple node names can be specified, separated by commas.
  Example: "ABC, bac"


  Configuration File Options:
  The following options can be used to filter node search results:
    - areaId: Only nodes from the specified area are considered
    - minRatedVoltage: Nodes with ratings below this value are excluded
    - nodeCharIndex and nodeChar: Skip nodes connecting generators to main nodes (e.g., YABC123)
    - changeValue: Value added to the connected node's voltage
    - skipFakeNodes: Exclude nodes ending with 55, which only exist in plans models
    - skipGeneratorsConnectedToNodesTypeOne: Exclude generators connected to nodes of type 1
    - skipGeneratorsWithoutTransformers: Exclude generators without directly connected transformers
*/

// Create a file system object for file operations
var fso = new ActiveXObject( "Scripting.FileSystemObject" );

// Initialize the configuration object using the specified file location
var config = iniConfigConstructor( configurationFileLocation, fso );
var tmpFile = config.homeFolder + "\\tmp.bin", tmpOgFile = config.homeFolder + "\\tmpOrg.bin";

// Load the KDM model file and save it as a temporary binary file
loadModel( config.modelPath, config.modelName );
if( SaveTempBIN( tmpOgFile ) < 1 ) errorThrower( "Unable to create temporary file" );

// Record the macro's start time
var time = getTime();

// Initialize arrays
var nodes = [], elements = [], baseElementsReactPow = [], baseNodesVolt = [];

// Set power flow calculation settings using the configuration file
setPowerFlowSettings( config );

// Calculate power flow; throw an error if the calculation fails
CPF();

// Try to read file from location specified in configuration file, then make array from file and close that file
var inputFile = readFile( config, fso );
var inputArray = getInputArray( inputFile );
inputFile.close();

// Create result files and folder using settings from a config file
var resultFiles = [ createFile( "info", "txt", config, fso ), createFile( "areas", "csv", config, fso ), 
createFile( "nodes", "csv", config, fso ), createFile( "generators", "csv", config, fso ), 
createFile( "transformers", "csv", config, fso ), createFile( "q", "csv", config, fso ), 
createFile( "v", "csv", config, fso ) ];
  
// Fills model info file ( resultFiles[ 0 ] ) with model info from configuration file and plans  
saveModelInfoToFile( resultFiles[ 0 ], config );

// Fills area file ( resultFiles[ 1 ] ) with areas from the model
saveAreasToFile( resultFiles[ 1 ], config );

// Fills node file ( resultFiles[ 2 ] ) with valid nodes
// Also fills nodes array with nodes that were written to the file and 
// baseNodesVolt with the base voltage of the nodes
fillNodesArrays(  nodes, baseNodesVolt, inputArray, resultFiles[ 2 ], config );

// Fills generator file ( resultFiles[ 3 ] ) with valid generators and their connected nodes
// Add valid generators to arrays with coresponding node and
// baseElementsReactPow with generators reactive power
fillGeneratorsArrays( elements, baseElementsReactPow, inputArray, resultFiles[ 3 ], config );

// Fills transformer file ( resultFiles[ 4 ] ) with valid transformers and connected nodes
// Add valid transformers to arrays with coresponding node and branch and
// baseElementsReactPow with transformers reactive power
fillTransformersArrays( elements, nodes, baseElementsReactPow, resultFiles[ 4 ], config );

// Writes to file headers and elements name
writeDataToFile( resultFiles[ 5 ], elements );

// Writes to file headers and nodes name
writeDataToFile( resultFiles[ 6 ], nodes );

// Trying to save file before any change on transformers and connected nodes
if( SaveTempBIN( tmpFile ) < 1 ) errorThrower( "Unable to create temporary file" );

// Initialize variables
var element = node = elementBaseValue = difference = buffer = elementType = null;

// For each element make some change depending on type of elemenet, then write results into result files
for( i in elements ){ 
  
  // Get the element and the connected node
  element = elements[ i ][ 0 ], node = elements[ i ][ 1 ];
  
  // If array element have a branch then try to switch tap up. If transformer is on it's last tap then change it down. 
  if( elements[ i ][ 2 ] ){

    elementType = "Unknown Transformer";

    // Write transformer's type by its number
    switch( element.Typ ){

      case 11: elementType = "wo regulation Transformer"; break;
      case 12: elementType = "w U-regulation Transformer"; break;
      case 13: elementType = "w P-regulation Transformer"; break;
      case 14: elementType = "Block Transformer"; break;
      case 15: elementType = "w P-regulation symetric Transformer"; break;
    }
    
    elementBaseValue = element.Stp0;

    if( ( element.TapLoc === 1 && element.Stp0 < element.Lstp ) || ( element.TapLoc === 0 && element.Stp0 <= 1 ) ) element.Stp0++;
    
    else element.Stp0--;  
    
    // Calculate power flow, if fails try to load original model and throw error 
    sCPF( tmpOgFile );

    difference = element.Stp0 - elementBaseValue;
  }

  // Get set value from config file and add it to node's voltage  
  else{ 
    
    elementType = "Unknown Generator";

    // Write generator's type by its number
    switch( element.Typ ){

      case 2: elementType = "JWCDc Generator"; break;
      case 3: elementType = "JWCDp Generator"; break;
      case 4: elementType = "JWCKc Generator"; break;
      case 5: elementType = "JWCKw Generator"; break;
      case 6: elementType = "JWCKp Generator"; break;
      case 7: elementType = "MC Generator"; break;
      case 8: elementType = "PR Generator"; break;
      case 9: elementType = "MW Generator"; break;
      case 10: elementType = "MF Generator"; break;
      case 11: elementType = "Other Generator"; break;
      case 12: elementType = "JWCKf Generator"; break;
      case 13: elementType = "PV Generator"; break;
      case 14: elementType = "JWCKpv Generator"; break;
    }

    elementBaseValue = roundTo( node.Vs, config.roundingPrecision );
    
    node.Vs += config.changeValue;

    // Calculate power flow, if fails try to load original model and throw error 
    sCPF( tmpOgFile );

    difference = roundTo( node.Vs - elementBaseValue, config.roundingPrecision );
  }
  
  // Write element's name, it's difference of connected node power / tap number to base
  buffer = strip( element.Name ) + ";" + difference + ";" + elementType + ";";

  // Write for each element it's new reactive power
  for( j in elements ){

    // Check if element have a branch, if true use reactive power from matching branch end
    react = elements[ j ][ 2 ] ? ( elements[ j ][ 0 ].begName === elements[ j ][ 1 ].Name ? elements[ j ][ 2 ].Qbeg : elements[ j ][ 2 ].Qend ) : elements[ j ][ 0 ].Qg;
  
    buffer += roundTo( react - baseElementsReactPow[ j ], config.roundingPrecision ) + ";";
  }
  
  resultFiles[ 5 ].WriteLine( removeLastChar( buffer ) );

  // Write element's name, it's difference of connected node power / tap number to base
  buffer = strip( element.Name ) + ";" + difference + ";" + elementType + ";";

  // Write for each node it's new voltage
  for( j in nodes ) buffer += roundTo( nodes[ j ].Vi - baseNodesVolt[ j ], config.roundingPrecision ) + ";";
  resultFiles[ 6 ].WriteLine( removeLastChar( buffer ) );

  // Load model without any changes to transformators
  ReadTempBIN( tmpFile );
}
// Load the original model from temporary backup
ReadTempBIN( tmpOgFile );

// Remove temporary binary files to clean up
fso.DeleteFile( tmpFile );
fso.DeleteFile( tmpOgFile );

// Close the result files to release resources
resultFiles[5].Close();
resultFiles[6].Close();

// Calculate and display the program's working duration.
// Note: The working duration calculation does not account for cases where the program 
// starts and ends on different days due to the lack of date handling.
var duration = getTime() - time;
MsgBox( "Task completed in " + formatTime( duration ), 0 | 64, "Task completed" );



// Loads the model from the specified path and name.
// The model path and name is split into a file extension, and the correct
// read function is called based on the file extension. 
function loadModel( modelPath, modelName ) {

  var model = modelPath + "\\" + modelName;

  // Get the file extension from the model name
  var fileExtension = strip( modelName.split( "." )[ 1].toLowerCase() );
 
  // Call the correct read function based on the file extension
  switch( fileExtension ){

    case "kdm": ReadDataKDM( model ); break;
    case "uzp": ReadDataUZP( model ); break;
    case "ien": ReadDataIEN( model ); break;
    case "bin": ReadDataBIN( model ); break;
    case "epc": ReadDataEPC( model ); break;
    case "pti": ReadDataPTI( model ); break;
    case "ucte": ReadDataUCTE( model ); break;
    default: errorThrower( "Unsupported file format" ); 
  }

}

//  Saves information about the model, including power flow calculation settings
//  to the specified file.
function saveModelInfoToFile( file, config ) {

  var format = method = "Unknown";

  // Determine the name of the power flow calculation method based on the
  // value of config.method
  switch( config.method ){

    case 1: method = "Stott"; break;
    case 2: method = "Newton"; break;
    case 3: method = "Ward"; break;
    case 4: method = "Gauss"; break;
    case 5: method = "DC"; break;
  }

  switch( Data.Format ){

    case 10: format = "KDM"; break;
    case 20: format = "UZP"; break;
    case 30: format = "IEN"; break;
    case 40: format = "BIN"; break;
    case 60: format = "EPC"; break;
    case 70: format = "PTI"; break;
    case 80: format = "UCTE"; break;
  }

  // Write the model information to the file
  file.Write(

    // Model information
    "Model: " + Data.FileName + "\n" +
    "Format: " + format + "\n" +
    "Date: " + getCurrentDate() + "\n\n" +
    // Power flow calculation settings
    "Power flow settings:\n" +
    "Method: " + method + "\n" +
    "Max iterations: " + config.maxIterations + "\n" +
    "Starting precision: " + config.startingPrecision + "\n" +
    "Precision: " + config.precision + "\n" +
    "Uzg iteration precision: " + config.uzgIterationPrecision + "\n\n" +
    // Variable settings
    "Variable settings:\n" +
    "Area: " + config.areaId + "\n" +
    "Min rated voltage: " + config.minRatedVoltage + "\n" +
    "Node char index: " + config.nodeCharIndex + "\n" +
    "Node char: " + config.nodeChar + "\n" +
    "Change value: " + config.changeValue + "\n" +
    "Skip fake nodes: " + config.skipFakeNodes + "\n" +
    "Skip generators connected to nodes type one: " + config.skipGeneratorsConnectedToNodesTypeOne + "\n" +
    "Skip generators without transformers: " + config.skipGeneratorsWithoutTransformers
  );

  // Close the file
  file.close();
}

// Saves all areas from the model to the specified file.
// The output file format is a CSV file with the following columns:
// Name, Area, Zone, Compound, Region
function saveAreasToFile( file, config ) {

  // Initialize a buffer string with headers
  var buffer = "name;area;zone;compound;region\n";
  var area = null;

  // Loop through all areas in the project
  for ( var i = 1; i < Data.N_Zon; i++ ){

    area = ZonArray.Get( i );

    // Check if the area matches the specified area or if the area index is 0
    if( area.Area === config.areaId || config.areaId <= 0 ){

      // Write the area name, area index, zone index, compound index, and region index to the buffer
      buffer += strip( area.AreaName ) + ";" + area.Area + ";" + area.Zone + ";" + area.Comp + ";" + area.Regn + "\n";
    }

  }

  // Write the buffer to the file and close it
  file.Write( buffer );
  file.close();
}

// Fills a node file with nodes from the specified area that meet certain criteria.
// The function also populates arrays with nodes that were written to the file and their base voltage.
function fillNodesArrays( nodesArray, baseNodesVoltageArray, inputArray, file, config ) {

  // Initialize a buffer string with headers
  var buffer = "name;type;min_voltage;current_voltage;max_voltage;area;zone;compound;region\n";
  var node = null;

  // Loop through all nodes in the project
  for (var i = 1; i < Data.N_Nod; i++) {

    node = NodArray.Get(i);

    // Skip nodes ending with 55 if skipFakeNodes is set in config
    if (config.skipFakeNodes && isStringMatchingRegex(strip(node.Name), "55$")) continue;

    // Add node to arrays if it fulfills all conditions
    if (
      (node.Area === config.areaId || config.areaId <= 0) && // Matches the specified area
      node.St > 0 && // Node is connected
      node.Name.charAt(config.nodeCharIndex) != config.nodeChar && // Not a generator's node
      node.Vn >= config.minRatedVoltage && // Voltage setpoint is higher than minimum
      isStringMatchingRegexArray(strip(node.Name), inputArray) // Name matches one from input
    ) {

      buffer += strip(node.Name) + ";" + node.Typ + ";"+ 
      roundTo(node.Vmin, config.roundingPrecision) + ";" +
      roundTo(node.Vi, config.roundingPrecision) + ";" + 
      roundTo(node.Vmax, config.roundingPrecision) + ";" +
      node.Area + ";" + node.Zone + ";" + node.Comp + ";" + node.Regn + "\n";

      nodesArray.push(node); // Add node to nodesArray
      baseNodesVoltageArray.push(node.Vi); // Add base voltage to baseNodesVoltageArray
    }
  }

  // Write buffered node data to file and close it
  file.Write(buffer);
  file.close();
}

// Fills the generator file with valid generators and their connected nodes.
// Also populates arrays with generators' reactive power and connected nodes' power.
function fillGeneratorsArrays( elementsArray, baseElementsReactPowerArray, inputArray, file, config ) {

  // Write headers to the file
  var buffer = "name;min_active_power;current_active_power;max_active_power;min_reactive_power;current_reactive_power;max_reactive_power;connected_node;zone\n";
  var element, node, branch;

  // Loop through all generators in the project
  for (var i = 1; i < Data.N_Gen; i++) {

    // Get generator and node that it's connected to
    element = GenArray.Get(i), node = NodArray.Get(element.NrNod);

    // Skip generators connected to type 1 nodes if specified in config
    if (config.skipGeneratorsConnectedToNodesTypeOne && node.Typ === 1) continue;

    // If generator is connected to a transformer, get the transformer and its connected node
    if (element.TrfName) {
    
      branch = BraArray.Find(element.TrfName);
      node = (node.Name === branch.EndName) ? NodArray.Find(branch.BegName) : NodArray.Find(branch.EndName);
    } 
    
    // Skip generators without transformers if specified in config
    else if (config.skipGeneratorsWithoutTransformers) continue;

    // Add valid generators to arrays based on constraints
    if (
      element.Qmin < element.Qmax && // Reactive power constraints
      element.St > 0 && // Generator is connected 
      ( node.Area === config.areaId || config.areaId <= 0 ) && // Area constraint
      isStringMatchingRegexArray( strip( node.Name ), inputArray ) // Name matching
    ) {
     
      // Write generator information to file
      buffer += strip( element.Name ) + ";" + 
      roundTo( element.Pmin, config.roundingPrecision ) + ";" + 
      roundTo( element.Pg, config.roundingPrecision ) + ";" + 
      roundTo( element.Pmax, config.roundingPrecision ) + ";" + 
      roundTo( element.Qmin, config.roundingPrecision ) + ";" + 
      roundTo( element.Qg, config.roundingPrecision ) + ";" + 
      roundTo( element.Qmax, config.roundingPrecision ) + ";" + 
      strip( node.Name ) + ";" +
      node.Zone + "\n";
   
      // Add generator to elements array and set its reactive power in baseElementsReactPowerArray
      elementsArray.push( [ element, node ] );
      baseElementsReactPowerArray.push( element.Qg );  
    }
  }

  // Write the buffer string to the file
  file.Write(buffer);
  file.close();
}

// Adds valid transformers to arrays with coresponding node and branch.
// The constraints are:
// Transformer must be connected to node from nodes array,
// have more than 1 tap, not already in elements array
function fillTransformersArrays( elements, nodes, baseElementsReactPowerArray, file, config ) {

  // Write headers to the file
  file.WriteLine( "name;min_tap;current_tap;max_tap;regulation_step;begining_node;ending_node;zone" );
  var node, transformer, branch;

  // Loop through all nodes in the nodes array
  for( i in nodes ){

    node = nodes[ i ];

    // Loop through all transformers in the project
    for( var j = 1; j < Data.N_Trf; j++ ){

      transformer = TrfArray.Get( j );
     
      if( 
        // Check if the transformer is connected to the node from nodes array
        ( node.Name == transformer.EndName || node.Name == transformer.BegName ) &&
        // Check if the transformer's node name contains the specified character
        transformer.EndName.charAt( config.nodeCharIndex ) != config.nodeChar && 
        // Check if the transformer's node name contains the specified character
        transformer.BegName.charAt( config.nodeCharIndex )!= config.nodeChar && 
        // Check if the transformer is not already in the elements array
        !isElementInArrayByName( elements, transformer.Name ) && transformer.Lstp > 1 && 
        // Check if the transformer's rated voltages are higher than the specified minimum
        transformer.Vn2 >= config.minRatedVoltage && transformer.Vn1 >= config.minRatedVoltage
      ){

        // If transformer is on it's last tap then minTap is set to 1 and maxTap is set to number of taps
        // If not then minTap is set to number of taps and maxTap is set to 1
        var minTap = maxTap = null;

        // Set minTap and maxTap based on the transformer's tap position
        if( transformer.TapLoc === 1 ){ minTap = 1, maxTap = transformer.Lstp; }

        else{ minTap = transformer.Lstp, maxTap = 1 }

        // Write to file transformer's name, current tap, min tap, max tap, tap regulation step and connected node
        file.WriteLine( strip( transformer.name ) + ";" + minTap + ";" + transformer.Stp0 + ";" + 
        maxTap + ";" + roundTo( transformer.dUstp, config.roundingPrecision ) + ";" + 
        strip( transformer.BegName ) + ";" + strip( transformer.EndName ) + ";" + node.Zone );
        
        // Get the branch connected to the transformer
        branch = BraArray.Get( transformer.NrBra );

        // Add the transformer to the elements array and set its reactive power in the baseElementsReactPowerArray
        elements.push( [ transformer, node, branch ] );
        
        // Set the reactive power based on the connected node
        if( node.Name === transformer.EndName ) baseElementsReactPowerArray.push( branch.Qend );
        
        else baseElementsReactPowerArray.push( branch.Qbeg );
      }

    }

  }

  file.close();
}

// Rounds a value to the specified number of decimal places.
// Uses the JavaScript Math.round() function.
function roundTo( value, precision ) {

  // Calculate the multiplier based on the precision
  var multiplier = Math.pow(10, precision);

  // Round the value to the specified number of decimal places
  return Math.round( value * multiplier ) / multiplier;
}

// Sets power flow calculation settings using the specified configuration object.
function setPowerFlowSettings( config ) {

  // Set the maximum number of iterations for the power flow calculation
  Calc.Itmax = config.maxIterations;

  // Set the starting precision for the power flow calculation
  Calc.EPS10 = config.startingPrecision;

  // Set the precision for the power flow calculation
  Calc.Eps = config.precision;

  // Set the precision for the Uzg iteration in the power flow calculation
  Calc.EpsUg = config.uzgIterationPrecision;

  // Set the power flow calculation method
  Calc.Met = config.method;
}

// Attempts to load the original file before throwing an error.
function safeErrorThrower( message, binPath ) {
  
  // Attempt to read the temporary binary file
  try { ReadTempBIN(binPath); } 
  
  // Show a message box if unable to load the original model
  catch (e) { MsgBox("Couldn't load original model", 16); }

  // Throw an error with the provided message
  errorThrower(message);
}

// Basic error thrower with error message window
// Shows error message in a message box and then throws the error
// Takes only one argument - error message
function errorThrower( message ) {
  
  MsgBox( message, 16, "Error" );

  throw new Error( message );
}

// Calls built-in power flow calculation function, throws error if it fails
function CPF() {

  if (CalcLF() !== 1) throw new Error("Power flow calculation failed");
}

// Calls built-in power flow calculation function, throws error if it fails, then tries to load original model
function sCPF( backupFile ) {

  // Call built-in power flow calculation function
  if( CalcLF() !== 1 ) {
    
    // If power flow calculation fails, try to load original file and throw error
    safeErrorThrower( "Power Flow calculation failed", backupFile );
  }
}

// Function takes string and returns it without whitespaces
function strip( string ) {

  // Remove leading and trailing whitespaces
  return string.replace(/(^\s+|\s+$)/g, '');
}

// Function checks if string matches regex and returns true/false
// Takes two arguments: string - string to check, regex - regex to check against
// Returns true if string matches regex, false if not
function isStringMatchingRegex( string, regex ) {

  // Regex flags for Multiline and Insensitive
  var regexExpression = new RegExp( regex , "mi" );  
    
  // Check if string matches regex
  if( string.search( regexExpression ) > -1 ) return true;
  
  return false;
}

// Function checks if string matches any of regexes in an array and returns true/false
// Takes two arguments: string - string to check, regexArray - array of regexes to check against
// Returns true if string matches any of regexes, false if not
function isStringMatchingRegexArray( string, regexArray ) {

  // Loop through each regex in the array and check if string matches it
  for( var i in regexArray ) if( isStringMatchingRegex( string, "^" + regexArray[ i ] ) ) return true;

  // If no regex matches, return false
  return false;
}

// Function takes a 2D array and an element name.
// It iterates through the array and checks if the name of any element matches the given element name.
function isElementInArrayByName( array, elementName ) {

  // Loop through each element in the array and check if it matches the given elementName
  for( i in array ) if(array[ i ][ 0 ].Name === elementName) return true;
  
  // If no element is found, return false
  return false;
}

// Function removes the last character from a string
function removeLastChar( string ) {

  return string.slice( 0, -1 );
}

// Takes two arguments: file - file object, objectArray - array of objects
// Writes to file headers, objects names and corresponding data
function writeDataToFile( file, objectArray ) {

  // Create buffer string with column names
  var buffer = "Elements;U_G/Tap Difference;Element Type;";

  // Loop through each object in the array and add it's name to the buffer string
  for( i in objectArray ){
  
    if( objectArray[ i ][ 0 ] ) buffer += strip( objectArray[ i ][ 0 ].Name ) + ";";

    else buffer += strip( objectArray[ i ].Name ) + ";";
  
  }

  // Write buffer string without the last character ( which is a comma ) to the file
  file.WriteLine( buffer.slice( 0, -1 ) );
}

// Creates a folder in the specified location based on the configuration object.
// Throws an error if the configuration object is null or if the folder cannot be created.
function createFolder( config, fso ) {

  // Check if the configuration object is null
  if (!config) errorThrower("Unable to load configuration");

  // Add timestamp to directory name if specified in configuration file
  var timeStamp = ( config.addTimestampToResultsDirectory == 1 ) ? getCurrentDate() + "--" : "";

  // Get folder name and path from the configuration
  var folder = timeStamp + config.folderName;
  var folderPath = config.homeFolder + "\\" + folder;

  // Check if the folder does not exist, then create it
  if (!fso.FolderExists(folderPath)) {
    
    try { fso.CreateFolder(folderPath); } 
    
    catch (e) { errorThrower("Unable to create folder"); }
  }

  // Return the folder path with a trailing backslash
  return folder + "\\";
}

// Function takes config object and depending on it's config creates file in specified location.
// Also can create folder where results are located depending on configuration file 
// Throws error if config object is null and when file can't be created
function createFile( name, format, config, fso ) {
 
  if( !config ) errorThrower( "Unable to load configuration" );

  var file = null;
  
  // Create folder if specified in configuration file
  var folder = ( config.createResultsFolder == 1 ) ? createFolder( config, fso ) : "";
  
  // Create file location path
  var fileLocation = config.homeFolder + "\\" + folder + name + "." +format;
  try{ file = fso.CreateTextFile( fileLocation ); }
  
  catch( e ){ errorThrower( "File already exists or unable to create it" ); }

  return file;
}

// Reads a file from the specified location based on the configuration object.
// Throws an error if the configuration object is null or if the file can't be read.
function readFile( config, fso ) {

  // Check if the configuration object is null
  if( !config ) errorThrower( "Unable to load configuration" );

  var file = null;

  // Get the file location from the configuration
  var fileLocation = config.inputFileLocation + "\\" + config.inputFileName;
    
  try{ 
    // Try to open the file
    file = fso.OpenTextFile( fileLocation, 1, false, 0 ); 
  } 
  
  catch( e ){ 
    // Throw an error if the file can't be opened
    errorThrower( "Unable to find or open file" ); 
  }

  return file;
}

// Function uses built in .ini function to get it's settings from config file.
// Returns conf object with settings taken from file. If file isn't found error is throwed instead.
function iniConfigConstructor( iniPath, fso ) {
  
  var configFile = iniPath + "\\config.ini";

  if( !fso.FileExists( configFile ) ) errorThrower( "config.ini file not found" );

  // Initializing plans built in ini manager
  var ini = CreateIniObject();
  ini.Open( configFile );

  var hFolder = ini.GetString( "main", "homeFolder", Main.WorkDir );
  
  // Declaring conf object and trying to fill it with config.ini configuration
  var config = {
  
    // Main
    homeFolder: hFolder,
    modelName: ini.GetString( "main", "modelName", "model" ),
    modelPath: ini.GetString( "main", "modelPath", hFolder ),  
   
    // Variable
    areaId: ini.GetInt( "variable", "areaId", 1 ),
    minRatedVoltage: ini.GetInt( "variable", "minRatedVoltage", 0 ),
    nodeCharIndex: ini.GetInt( "variable", "nodeCharIndex", 0 ),
    nodeChar: ini.GetString( "variable", "nodeChar", 'Y' ),
    changeValue: ini.GetInt( "variable", "changeValue", 1 ),
    skipFakeNodes: ini.GetBool( "variable", "skipFakeNodes", 0 ),
    skipGeneratorsConnectedToNodesTypeOne: ini.GetBool( "variable", "skipGeneratorsConnectedToNodesTypeOne", 0 ),
    skipGeneratorsWithoutTransformers: ini.GetBool( "variable", "skipGeneratorsWithoutTransformers", 0 ),
    
    // Folder
    createResultsFolder: ini.GetBool( "folder", "createResultsFolder", 0 ),
    folderName: ini.GetString( "folder", "folderName", "folder" ),
    addTimestampToResultsDirectory: ini.GetBool( "files", "addTimestampToResultsDirectory", 1 ),

    // Files
    inputFileLocation: ini.GetString( "files", "inputFileLocation", hFolder ),
    inputFileName: ini.GetString( "files", "inputFileName", "input" ),
    roundingPrecision: ini.GetInt( "files", "roundingPrecision", 2 ),
    
    // Power Flow
    maxIterations: ini.GetInt( "power flow", "maxIterations", 300 ),
    startingPrecision: ini.GetDouble( "power flow", "startingPrecision", 10.00 ),
    precision: ini.GetDouble( "power flow", "precision", 1.00 ),
    uzgIterationPrecision: ini.GetDouble( "power flow", "uzgIterationPrecision", 0.001 ),
    method: ini.GetInt( "power flow", "method", 1 )
  };
  
  // Overwriting config.ini file

  // Main
  ini.WriteString( "main", "homeFolder", config.homeFolder );
  ini.WriteString( "main", "modelName", config.modelName );
  ini.WriteString( "main", "modelPath", config.modelPath );
  
  // Variable
  ini.WriteInt( "variable", "areaId", config.areaId );
  ini.WriteInt( "variable", "minRatedVoltage", config.minRatedVoltage );
  ini.WriteInt( "variable", "nodeCharIndex", config.nodeCharIndex );
  ini.WriteString( "variable", "nodeChar", config.nodeChar );
  ini.WriteInt( "variable", "changeValue", config.changeValue );
  ini.WriteBool( "variable", "skipFakeNodes", config.skipFakeNodes );
  ini.WriteBool( "variable", "skipGeneratorsConnectedToNodesTypeOne", config.skipGeneratorsConnectedToNodesTypeOne );
  ini.WriteBool( "variable", "skipGeneratorsWithoutTransformers", config.skipGeneratorsWithoutTransformers );
  
  // Folder
  ini.WriteBool( "folder", "createResultsFolder", config.createResultsFolder );
  ini.WriteString( "folder", "folderName", config.folderName );
  ini.WriteBool( "files", "addTimestampToResultsDirectory", config.addTimestampToResultsDirectory );
    
  // Files
  ini.WriteString( "files", "inputFileLocation", config.inputFileLocation );
  ini.WriteString( "files", "inputFileName", config.inputFileName );
  ini.WriteInt( "files", "roundingPrecision", config.roundingPrecision );
    
  // Power Flow
  ini.WriteInt( "power flow", "maxIterations", config.maxIterations );
  ini.WriteDouble( "power flow", "startingPrecision", config.startingPrecision );
  ini.WriteDouble( "power flow", "precision", config.precision );
  ini.WriteDouble( "power flow", "uzgIterationPrecision", config.uzgIterationPrecision );
  ini.WriteInt( "power flow", "method", config.method );
 
  return config;
}

// Function gets file and takes each line into a array and after finding "," character pushes array into other array. 
// Returns string array
function getInputArray( file ) {

  // Initialize temporary array and word variable
  var array = [], tmp = [], word;

  // Loop through file until its end
  while( !file.AtEndOfStream ){

    // Split line by commas and loop through it
    tmp = file.ReadLine().split(",");

    for( i in tmp ){

      // Remove leading and trailing spaces from the word
      word = tmp[ i ].replace(/(^\s+|\s+$)/g, '');

      // If word isn't empty push it to array
      if( word ) array.push( word );
    }

  }

  return array;
}

// Function takes current date and returns it in file safe format  
function getCurrentDate() {
  
  var current = new Date()
  
  var formatedDateArray = [ ( '0' + ( current.getUTCMonth() + 1 ) ).slice( -2 ), ( '0' + current.getUTCDate() ).slice( -2 ), 
  ( '0' + current.getUTCHours() ).slice( -2 ), ( '0' + current.getUTCMinutes() ).slice( -2 ), ( '0' + current.getUTCSeconds() ).slice( -2 ) ];
  
  return current.getUTCFullYear() + "-" + formatedDateArray[ 0 ] + "-" + formatedDateArray[ 1 ] + "--" + formatedDateArray[ 2 ] + "-" + formatedDateArray[ 3 ] + "-" + formatedDateArray[ 4 ];
}

// Function returns current time UTC in seconds 
function getTime() {
  
  var current = new Date();
  
  return current.getHours() * 3600 + current.getMinutes() * 60 + current.getSeconds();
}

// Function takes time in seconds and returns time in HH:MM:SS format
function formatTime( time ) {

  var hours = minutes = 0;

  hours = Math.floor( time / 3600 );

  time -= hours * 3600;

  minutes = Math.floor( time / 60 );

  time -= minutes * 60;

  return ( '0' + hours ).slice( -2 ) + ":" + ( '0' + minutes ).slice( -2 ) + ":" + ( '0' + time ).slice( -2 );
}
