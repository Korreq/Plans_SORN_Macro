# Plans_SORN_Macro
Macro gets nodes' name from input file, then searches through nodes to find connected transformers and generators ( directly or through Y - node ).
Before any changes original values for each nodes and elements are written with their names into both output files. 
First file is for voltage change results, second is for reactive power changes.    
Then for each found generator, increases it's set voltage by value from configuration file and 
writes it's new value with new calculated values of each element/node into files.       
Similarly for each transformer except changeing tap by one to increase it's voltage. 
After succesfully writing results into files, macro shows in a message window time it took to be done.