# Plans_SORN_Macro
Description:

    Macro gets nodes' name from input file, then searches through nodes to find connected transformers and generators ( directly or through Y - node ).
    Before any changes original values for each nodes and elements are written with their names into both output files. 
    First file is for voltage change results, second is for reactive power changes.    
    Then for each found generator, increases it's set voltage by value from configuration file and 
    writes it's new value with new calculated values of each element/node into files.       
    Similarly for each transformer except changeing tap by one to increase it's voltage. 
    After succesfully writing results into files, macro shows in a message window time it took to be done.

Writing nodes names in an input file:
  
    Nodes are searched by function using regex. It matches input from start of string and is case-insensitive.
    E.g. we have these nodes: ABC111, ABC222, ABC555, BAC234, BAC235, BAC124, ABCA123, CABC345, ZABC111.
    If we want all nodes containing name "ABC" and "BAC" then in input file we can write: 
  
    ABC, bac
  
    White spaces dosen't matter, macro splits input by ',' if we forget to add ',' between searched strings:
  
    ABC bac 
  
    Macro will search for nodes containing "ABCBAC".
    If we want to find nodes ABC111, ABC222 and only them, then we can write:
  
    ABC111, ABC222
  
    Macro will only find ABC111, ABC222 nodes.
  
    Thakns to regex support we are able to find more specific nodes. E.g. we want to find all BAC nodes ending with 4:
  
    bac..4
  
    Will find nodes BAC234 and BAC124.
  
    In configuration file there are options that can block finding some nodes:
  
      areaId - nodes only from this area will be found, 
      minRatedVoltage - nodes that are rated less than specified will not be found,
      nodeCharIndex and nodeChar - this options are for skiping nodes connecting generators to main nodes e.g. YABC123,
      skipFakeNodes - nodes that end on 55 which due to model implementation in plans don't have real representation

