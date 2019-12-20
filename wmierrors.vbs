'-- Partial listing here derived from WMI diagnostic utility: http://www.microsoft.com/en-us/download/details.aspx?id=7684
'-- But for some reason that utility had less complete list of error codes than some other sources.

'Dim shex, sDesc
'sDesc = GetWMIErrDesc( -2147217397, shex)
'MsgBox shex & vbCrLf & sDesc


Function GetWMIErrDesc(ErrNumber, ErrHex)
    On Error Resume Next
            ErrHex = UCase(Hex(ErrNumber))
         Select Case ErrNumber
        Case -2147221020  '-- 800401E4
                    GetWMIErrDesc = "'GetObject - object not available"
        Case -2147217407  '-- 80041001
                     GetWMIErrDesc = "Call failed"
        Case -2147217406  '-- 80041002
                     GetWMIErrDesc = "Object not found"
        Case -2147217405  '-- 80041003
                     GetWMIErrDesc = "Access denied"
        Case -2147217404  '-- 80041004
                     GetWMIErrDesc = "Provider has failed at some time other than during initialization"
        Case -2147217403  '-- 80041005
                     GetWMIErrDesc = "Type mismatch"
        Case -2147217402  '-- 80041006
                     GetWMIErrDesc = "Out of memory"
        Case -2147217401  '-- 80041007
                     GetWMIErrDesc = "The IWbemContext object is not valid"
        Case -2147217400  '-- 80041008
                     GetWMIErrDesc = "Invalid parameter"
        Case -2147217399  '-- 80041009
                     GetWMIErrDesc = "Resource typically a remote server is not currently available" 
        Case -2147217398  '-- 8004100A
                     GetWMIErrDesc = "Internal critical and unexpected error occurred." 
        Case -2147217397  '-- 8004100B
                     GetWMIErrDesc = "One or more network packets were corrupted during a remote session"
        Case -2147217396  '-- 8004100C
                     GetWMIErrDesc = "Not supported"
        Case -2147217395  '-- 8004100D
                     GetWMIErrDesc = "Parent class specified is not valid"
        Case -2147217394  '-- 8004100E
                     GetWMIErrDesc = "Invalid namespace"
        Case -2147217393  '-- 8004100F
                     GetWMIErrDesc = "Invalid object"
        Case -2147217392  '-- 80041010
                     GetWMIErrDesc = "Invalid class"
        Case -2147217391  '-- 80041011
                     GetWMIErrDesc = "Provider not found"
        Case -2147217390  '-- 80041012
                     GetWMIErrDesc = "Invalid provider registration"
        Case -2147217389  '-- 80041013
                     GetWMIErrDesc = "Provider load failed"
        Case -2147217388  '-- 80041014
                     GetWMIErrDesc = "Provider initialize failure" 
        Case -2147217387  '-- 80041015
                     GetWMIErrDesc = "Networking error"
        Case -2147217386  '-- 80041016
                     GetWMIErrDesc = "Invalid operation. This error usually applies to invalid attempts to delete classes or properties"
        Case -2147217385  '-- 80041017
                     GetWMIErrDesc = "Invalid query"
        Case -2147217384  '-- 80041018
                     GetWMIErrDesc = "Invalid query language"
        Case -2147217383  '-- 80041019
                     GetWMIErrDesc = "In a put operation the wbemChangeFlagCreateOnly flag was specified but the instance already exists" 
        Case -2147217382  '-- 8004101A
                     GetWMIErrDesc = "Not possible to perform the add operation on this qualifier because the owning object does not permit overrides"
        Case -2147217381  '-- 8004101B
                     GetWMIErrDesc = "Attempted to delete a qualifier that was not owned. The qualifier was inherited from a parent class"
        Case -2147217380  '-- 8004101C
                     GetWMIErrDesc = "Attempted to delete a property that was not owned. The property was inherited from a parent class"
        Case -2147217379  '-- 8004101D
                     GetWMIErrDesc = "Unexpected or illegal sequence of calls"
        Case -2147217378  '-- 8004101E
                     GetWMIErrDesc = "Illegal operation"
        Case -2147217377  '-- 8004101F
                     GetWMIErrDesc = "Illegal attempt to specify a key qualifier on a property that cannot be a key"
        Case -2147217376  '-- 80041020
                     GetWMIErrDesc = "Current object is not a valid class definition. Either it is incomplete or it has not been registered"
        Case -2147217375  '-- 80041021
                     GetWMIErrDesc = "Invalid query syntax"
       ' Case -2147217374  '-- 80041022  Reserved for future use
                    
        Case -2147217373  '-- 80041023
                     GetWMIErrDesc = "An attempt was made to modify a read-only property"
        Case -2147217372  '-- 80041024
                     GetWMIErrDesc = "Provider cannot perform the requested operation"
        Case -2147217371  '-- 80041025
                     GetWMIErrDesc = "Attempt was made to make a change that invalidates a subclass"
        Case -2147217370  '-- 80041026
                     GetWMIErrDesc = "Attempt was made to delete or modify a class that has instances"
       ' Case -2147217369  '-- 80041027   Reserved for future use
        
         Case -2147217368  '-- 80041028
                     GetWMIErrDesc = "Value of Nothing/NULL was specified for a property that must have a value"
        Case -2147217367  '-- 80041029
                     GetWMIErrDesc = "Invalid qualifier type"
        Case -2147217366  '-- 8004102A
                     GetWMIErrDesc = "Invalid property type"
        Case -2147217365  '-- 8004102B
                     GetWMIErrDesc = "Out-of-range or incompatible with the type"
        Case -2147217364  '-- 8004102C
                     GetWMIErrDesc = "Illegal attempt was made to make a class singleton such as when the class is derived from a non-singleton class" 
        Case -2147217363  '-- 8004102D
                     GetWMIErrDesc = "CIM type specified is invalid"
        Case -2147217362  '-- 8004102E
                     GetWMIErrDesc = "Requested method is not available"
        Case -2147217361  '-- 8004102F
                     GetWMIErrDesc = "Invalid parameter"
        Case -2147217360  '-- 80041030
                     GetWMIErrDesc = "Attempting to get qualifiers on a system property"
        Case -2147217359  '-- 80041031
                     GetWMIErrDesc = "Invalid property type"
        Case -2147217358  '-- 80041032
                     GetWMIErrDesc = "Asynchronous process has been canceled"
        Case -2147217357  '-- 80041033
                     GetWMIErrDesc = "Operation requested while WMI is shutting down"
        Case -2147217356  '-- 80041034
                     GetWMIErrDesc = "Attempt was made to reuse an existing method name from a parent class and the signatures do not match"
        Case -2147217355  '-- 80041035
                     GetWMIErrDesc = "One or more parameter values such as a query text is too complex or unsupported"
        Case -2147217354  '-- 80041036
                     GetWMIErrDesc = "Parameter was missing from the method call"
        Case -2147217353  '-- 80041037
                     GetWMIErrDesc = "Method parameter has an invalid ID qualifier"
        Case -2147217352  '-- 80041038
                     GetWMIErrDesc = "One or more of the method parameters have ID qualifiers that are out of sequence"
        Case -2147217351  '-- 80041039
                     GetWMIErrDesc = "Return value for a method has an ID qualifier"
        Case -2147217350  '-- 8004103A
                     GetWMIErrDesc = "Specified object path was invalid"
        Case -2147217349  '-- 8004103B
                     GetWMIErrDesc = "Disk is out of space or the 4 GB limit on WMI repository (WMI repository) size is reached"
        Case -2147217348  '-- 8004103C
                     GetWMIErrDesc = "Supplied buffer was too small to hold all of the objects in the enumerator or to read a string property"
        Case -2147217347  '-- 8004103D
                     GetWMIErrDesc = "Provider does not support the requested put operation"
        Case -2147217346  '-- 8004103E
                     GetWMIErrDesc = "Object with an incorrect type or version was encountered during marshaling"
        Case -2147217345  '-- 8004103F
                     GetWMIErrDesc = "Packet with an incorrect type or version was encountered during marshaling"
        Case -2147217344  '-- 80041040
                     GetWMIErrDesc = "Packet has an unsupported version"
        Case -2147217343  '-- 80041041
                     GetWMIErrDesc = "Packet appears to be corrupt"
        Case -2147217342  '-- 80041042
                     GetWMIErrDesc = "Attempt was made to mismatch qualifiers such as putting [key] on an object instead of a property" 
        Case -2147217341  '-- 80041043
                     GetWMIErrDesc = "Duplicate parameter was declared in a CIM method"
        Case -2147217340  '-- 80041044
                     GetWMIErrDesc = "Reserved for future use"
        Case -2147217339  '-- 80041045
                     GetWMIErrDesc = "Call to IWbemObjectSink::Indicate has failed. The provider can refire the event"
        Case -2147217338  '-- 80041046
                     GetWMIErrDesc = "Specified qualifier flavor was invalid"
        Case -2147217337  '-- 80041047
                     GetWMIErrDesc = "Attempt was made to create a reference that is circular (for example deriving a class from itself)" 
        Case -2147217336  '-- 80041048
                     GetWMIErrDesc = "Specified class is not supported"
        Case -2147217335  '-- 80041049
                     GetWMIErrDesc = "Attempt was made to change a key when instances or subclasses are already using the key"
        Case -2147217328  '-- 80041050
                     GetWMIErrDesc = "An attempt was made to change an index when instances or subclasses are already using the index"
        Case -2147217327  '-- 80041051
                     GetWMIErrDesc = "Attempt was made to create more properties than the current version of the class supports"
        Case -2147217326  '-- 80041052
                     GetWMIErrDesc = "Property was redefined with a conflicting type in a derived class"
        Case -2147217325  '-- 80041053
                     GetWMIErrDesc = "Attempt was made in a derived class to override a qualifier that cannot be overridden"
        Case -2147217324  '-- 80041054
                     GetWMIErrDesc = "Method was re-declared with a conflicting signature in a derived class"
        Case -2147217323  '-- 80041055
                     GetWMIErrDesc = "Attempt was made to execute a method not marked with [implemented] in any relevant class"
        Case -2147217322  '-- 80041056
                     GetWMIErrDesc = "Attempt was made to execute a method marked with [disabled]"
        Case -2147217321  '-- 80041057
                     GetWMIErrDesc = "Refresher is busy with another operation"
        Case -2147217320  '-- 80041058
                     GetWMIErrDesc = "Filtering query is syntactically invalid"
        Case -2147217319  '-- 80041059
                     GetWMIErrDesc = "The FROM clause of a filtering query references a class that is not an event class (not derived from __Event)"
        Case -2147217318  '-- 8004105A
                     GetWMIErrDesc = "A GROUP BY clause was used without the corresponding GROUP WITHIN clause"
        Case -2147217317  '-- 8004105B
                     GetWMIErrDesc = "A GROUP BY clause was used. Aggregation on all properties is not supported"
        Case -2147217316  '-- 8004105C
                     GetWMIErrDesc = "Dot notation was used on a property that is not an embedded object"
        Case -2147217315  '-- 8004105D
                     GetWMIErrDesc = "A GROUP BY clause references a property that is an embedded object without using dot notation"
        Case -2147217313  '-- 8004105F
                     GetWMIErrDesc = "Event provider registration query (__EventProviderRegistration) did not specify the classes for which events were provided"
        Case -2147217312  '-- 80041060
                     GetWMIErrDesc = "Request was made to back up or restore the WMI repository while it was in use by WinMgmt.exe or in Windows XP or later the SVCHOST process that contains the Windows Management service" 
        Case -2147217311  '-- 80041061
                     GetWMIErrDesc = "Asynchronous delivery queue overflowed from the event consumer being too slow"
        Case -2147217310  '-- 80041062
                     GetWMIErrDesc = "Operation failed because the client did not have the necessary security privilege"
        Case -2147217309  '-- 80041063
                     GetWMIErrDesc = "Operator is invalid for this property type"
        Case -2147217308  '-- 80041064
                     GetWMIErrDesc = "User specified a username/password/authority on a local connection. The user must use a blank username/password and rely on default security"
        Case -2147217307  '-- 80041065
                     GetWMIErrDesc = "Class was made abstract when its parent class is not abstract"
        Case -2147217306  '-- 80041066
                     GetWMIErrDesc = "Amended object was written without the WBEM_FLAG_USE_AMENDED_QUALIFIERS flag being specified"
        Case -2147217305  '-- 80041067
                     GetWMIErrDesc = "Client did not retrieve objects quickly enough from an enumeration"
        Case -2147217304  '-- 80041068
                     GetWMIErrDesc = "Null security descriptor was used"
        Case -2147217303  '-- 80041069
                     GetWMIErrDesc = "Operation timed out"
        Case -2147217302  '-- 8004106A
                     GetWMIErrDesc = "Association is invalid"
        Case -2147217301  '-- 8004106B
                     GetWMIErrDesc = "Operation was ambiguous"
        Case -2147217300  '-- 8004106C
                     GetWMIErrDesc = "WMI is taking up too much memory" 
        Case -2147217299  '-- 8004106D
                     GetWMIErrDesc = "Operation resulted in a transaction conflict"
        Case -2147217298  '-- 8004106E
                     GetWMIErrDesc = "Transaction forced a rollback"
        Case -2147217297  '-- 8004106F
                     GetWMIErrDesc = "Locale used in the call is not supported"
        Case -2147217296  '-- 80041070
                     GetWMIErrDesc = "Object handle is out-of-date"
        Case -2147217295  '-- 80041071
                     GetWMIErrDesc = "Connection to the SQL database failed"
        Case -2147217294  '-- 80041072
                     GetWMIErrDesc = "Handle request was invalid"
        Case -2147217293  '-- 80041073
                     GetWMIErrDesc = "Property name contains more than 255 characters"
        Case -2147217292  '-- 80041074
                     GetWMIErrDesc = "Class name contains more than 255 characters"
        Case -2147217291  '-- 80041075
                     GetWMIErrDesc = "Method name contains more than 255 characters"
        Case -2147217290  '-- 80041076
                     GetWMIErrDesc = "Qualifier name contains more than 255 characters"
        Case -2147217289  '-- 80041077
                     GetWMIErrDesc = "The SQL command must be rerun because there is a deadlock in SQL"
        Case -2147217288  '-- 80041078
                     GetWMIErrDesc = "Database version does not match the version that the WMI repository driver understands"
        Case -2147217287  '-- 80041079
                     GetWMIErrDesc = "WMI cannot execute the delete operation because the provider does not allow it"
        Case -2147217286  '-- 8004107A
                     GetWMIErrDesc = "WMI cannot execute the put operation because the provider does not allow it"
        Case -2147217280  '-- 80041080
                     GetWMIErrDesc = "Specified locale identifier was invalid for the operation"
        Case -2147217279  '-- 80041081
                     GetWMIErrDesc = "Provider is suspended"
        Case -2147217278  '-- 80041082
                     GetWMIErrDesc = "Object must be written to the WMI repository and retrieved again before the requested operation can succeed"
        Case -2147217277  '-- 80041083
                     GetWMIErrDesc = "Operation cannot be completed; no schema is available"
        Case -2147217276  '-- 80041084
                     GetWMIErrDesc = "Provider cannot be registered because it is already registered"
        Case -2147217275  '-- 80041085
                     GetWMIErrDesc = "Provider was not registered"
        Case -2147217274  '-- 80041086
                     GetWMIErrDesc = "Fatal transport error occurred"
        Case -2147217273  '-- 80041087
                     GetWMIErrDesc = "User attempted to set a computer name or domain without an encrypted connection"
        Case -2147217272  '-- 80041088
                     GetWMIErrDesc = "A provider failed to report results within the specified timeout"
        Case -2147217271  '-- 80041089
                     GetWMIErrDesc = "User attempted to put an instance with no defined key"
        Case -2147217270  '-- 8004108A
                     GetWMIErrDesc = "User attempted to register a provider instance but the COM server for the provider instance was unloaded"
        Case -2147213309  '-- 80042003
                     GetWMIErrDesc = "This computer does not have the necessary domain permissions" 
        Case -2147213311  '-- 80042001
                     GetWMIErrDesc = "Registration too broad"
        Case -2147213310  '-- 80042002
                     GetWMIErrDesc = "A WITHIN clause was not used in this query" 
                     
      '  Case -2147209215  '-- 80043001   "Reserved for future use" 
      '  Case -2147209214  '-- 80043002   "Reserved for future use" 

        Case -2147205119  '-- 80044001
             GetWMIErrDesc = "Expected a qualifier name."
        Case -2147205118  '-- 80044002
             GetWMIErrDesc = "Expected semicolon or '='."
        Case -2147205117  '-- 80044003
             GetWMIErrDesc = "Expected an opening brace."
         Case -2147205116  '-- 80044004
             GetWMIErrDesc = "Missing closing brace or an illegal array element."
         Case -2147205115  '-- 80044005
             GetWMIErrDesc = "Expected a closing bracket."
         Case -2147205114  '-- 80044006
             GetWMIErrDesc = "Expected closing parenthesis."
         Case -2147205113  '-- 80044007
             GetWMIErrDesc = "Numeric value out of range or strings without quotes."
         Case -2147205112  '-- 80044008
             GetWMIErrDesc = "Expected a type identifier."
         Case -2147205111  '-- 80044009
             GetWMIErrDesc = "Expected an open parenthesis."
         Case -2147205110  '-- 8004400A
             GetWMIErrDesc = "Unexpected token in the file."
         Case -2147205109  '-- 8004400B
             GetWMIErrDesc = "Unrecognized or unsupported type identifier."
         Case -2147205109  '-- 8004400B
             GetWMIErrDesc = "Expected property or method name."
         Case -2147205107  '-- 8004400D
             GetWMIErrDesc = "Typedefs and enumerated types are not supported."
         Case -2147205106  '-- 8004400E
             GetWMIErrDesc = "Only a reference to a class object can have an alias value."
         Case -2147205105  '-- 8004400F
             GetWMIErrDesc = "Unexpected array initialization. Arrays must be declared with []."
         Case -2147205104  '-- 80044010
             GetWMIErrDesc = "Namespace path syntax is not valid."
         Case -2147205103  '-- 80044011
             GetWMIErrDesc = "Duplicate amendment specifiers."
         Case -2147205102  '-- 80044012
             GetWMIErrDesc = "#pragma must be followed by a valid keyword."
         Case -2147205101  '-- 80044013
             GetWMIErrDesc = "Namespace path syntax is not valid."
         Case -2147205100  '-- 80044014
             GetWMIErrDesc = "Unexpected character in class name must be an identifier."
         Case -2147205099  '-- 80044015
             GetWMIErrDesc = "The value specified cannot be made into the appropriate type."
         Case -2147205098  '-- 80044016
             GetWMIErrDesc = "Dollar sign must be followed by an alias name as an identifier."
         Case -2147205097  '-- 80044017
             GetWMIErrDesc = "Class declaration is not valid."
         Case -2147205096  '-- 80044018
             GetWMIErrDesc = "The instance declaration is not valid. It must start with 'instance of'"
         Case -2147205095  '-- 80044019
             GetWMIErrDesc = "Expected dollar sign. An alias in the form '$name' must follow the 'as' keyword."
         Case -2147205094  '-- 8004401A
             GetWMIErrDesc = "'CIMTYPE' qualifier cannot be specified directly in a MOF file. Use standard type notation."
         Case -2147205093  '-- 8004401B
             GetWMIErrDesc = "Duplicate property name was found in the MOF."
         Case -2147205092  '-- 8004401C
             GetWMIErrDesc = "Namespace syntax is not valid. References to other servers are not allowed."
         Case -2147205091  '-- 8004401D
             GetWMIErrDesc = "Value out of range."
         Case -2147205090  '-- 8004401E
             GetWMIErrDesc = "The file is not a valid text MOF file or binary MOF file."
         Case -2147205089  '-- 8004401F
             GetWMIErrDesc = "Embedded objects cannot be aliases."
         Case -2147205088  '-- 80044020
             GetWMIErrDesc = "NULL elements in an array are not supported."
         Case -2147205087  '-- 80044021
             GetWMIErrDesc = "Qualifier was used more than once on the object."
         Case -2147205086  '-- 80044022
             GetWMIErrDesc = "Expected a flavor type such as ToInstance, ToSubClass, EnableOverride, or DisableOverride."
         Case -2147205085  '-- 80044023
             GetWMIErrDesc = "Combining EnableOverride and DisableOverride on same qualifier is not legal."
         Case -2147205084  '-- 80044024
             GetWMIErrDesc = "An alias cannot be used twice."
         Case -2147205083  '-- 80044025
             GetWMIErrDesc = "Combining Restricted, and ToInstance or ToSubClass is not legal."
         Case -2147205082  '-- 80044026
             GetWMIErrDesc = "Methods cannot return array values."
         Case -2147205081  '-- 80044027
             GetWMIErrDesc = "Arguments must have an In or Out qualifier."
         Case -2147205080  '-- 80044028
             GetWMIErrDesc = "Flags syntax is not valid."
         Case -2147205079  '-- 80044029
             GetWMIErrDesc = "The final brace and semi-colon for a class are missing."
         Case -2147205078  '-- 8004402A
             GetWMIErrDesc = "A CIM version 2.2 feature is not supported for a qualifier value."
         Case -2147205077  '-- 8004402B
             GetWMIErrDesc = "The CIM version 2.2 data type is not supported."
         Case -2147205076  '-- 8004402C
             GetWMIErrDesc = "Invalid DeleteInstance syntax"
         Case -2147205075  '-- 8004402D
             GetWMIErrDesc = "Invalid qualifier syntax"
         Case -2147205074  '-- 8004402E
             GetWMIErrDesc = "The qualifier is used outside of its scope."
         Case -2147205073  '-- 8004402F
             GetWMIErrDesc = "Error creating temporary file. The temporary file is an intermediate stage in the MOF compilation."
         Case -2147205072  '-- 80044030
             GetWMIErrDesc = "A file included in the MOF by the preprocessor command #include is not valid."
         Case -2147205071  '-- 80044031
             GetWMIErrDesc = "The syntax for the preprocessor commands #pragma deleteinstance or #pragma deleteclass is not valid."

         Case Else
                GetWMIErrDesc = "Not a WMI error code." 
     End Select

End Function