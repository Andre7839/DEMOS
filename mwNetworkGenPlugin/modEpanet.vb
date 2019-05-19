

Module modEpaNet

    'EPANET2.BAS
    '
    'Declarations of functions in the EPANET PROGRAMMERs TOOLKIT
    '(EPANET2.DLL)

    'Last updated on 4/3/07
    'ver last updated on 16/02/2011 'its

    ' These are codes used by the DLL functions
    Public Const EN_ELEVATION As Int16 = 0     ' Node parameters
    Public Const EN_BASEDEMAND As Int16 = 1
    Public Const EN_PATTERN As Int16 = 2
    Public Const EN_EMITTER As Int16 = 3
    Public Const EN_INITQUAL As Int16 = 4
    Public Const EN_SOURCEQUAL As Int16 = 5
    Public Const EN_SOURCEPAT As Int16 = 6
    Public Const EN_SOURCETYPE As Int16 = 7
    Public Const EN_TANKLEVEL As Int16 = 8
    Public Const EN_DEMAND As Int16 = 9
    Public Const EN_HEAD As Int16 = 10
    Public Const EN_PRESSURE As Int16 = 11
    Public Const EN_QUALITY As Int16 = 12
    Public Const EN_SOURCEMASS As Int16 = 13
    Public Const EN_INITVOLUME As Int16 = 14
    Public Const EN_MIXMODEL As Int16 = 15
    Public Const EN_MIXZONEVOL As Int16 = 16

    Public Const EN_TANKDIAM As Int16 = 17
    Public Const EN_MINVOLUME As Int16 = 18
    Public Const EN_VOLCURVE As Int16 = 19
    Public Const EN_MINLEVEL As Int16 = 20
    Public Const EN_MAXLEVEL As Int16 = 21
    Public Const EN_MIXFRACTION As Int16 = 22
    Public Const EN_TANK_KBULK As Int16 = 23

    Public Const EN_DIAMETER As Int16 = 0      ' Link parameters
    Public Const EN_LENGTH As Int16 = 1
    Public Const EN_ROUGHNESS As Int16 = 2
    Public Const EN_MINORLOSS As Int16 = 3
    Public Const EN_INITSTATUS As Int16 = 4
    Public Const EN_INITSETTING As Int16 = 5
    Public Const EN_KBULK As Int16 = 6
    Public Const EN_KWALL As Int16 = 7
    Public Const EN_FLOW As Int16 = 8
    Public Const EN_VELOCITY As Int16 = 9
    Public Const EN_HEADLOSS As Int16 = 10
    Public Const EN_STATUS As Int16 = 11
    Public Const EN_SETTING As Int16 = 12
    Public Const EN_ENERGY As Int16 = 13

    Public Const EN_DURATION As Int16 = 0      ' Time parameters
    Public Const EN_HYDSTEP As Int16 = 1
    Public Const EN_QUALSTEP As Int16 = 2
    Public Const EN_PATTERNSTEP As Int16 = 3
    Public Const EN_PATTERNSTART As Int16 = 4
    Public Const EN_REPORTSTEP As Int16 = 5
    Public Const EN_REPORTSTART As Int16 = 6
    Public Const EN_RULESTEP As Int16 = 7
    Public Const EN_STATISTIC As Int16 = 8
    Public Const EN_PERIODS As Int16 = 9

    Public Const EN_NODECOUNT As Int16 = 0     'Component counts
    Public Const EN_TANKCOUNT As Int16 = 1
    Public Const EN_LINKCOUNT As Int16 = 2
    Public Const EN_PATCOUNT As Int16 = 3
    Public Const EN_CURVECOUNT As Int16 = 4
    Public Const EN_CONTROLCOUNT As Int16 = 5

    Public Const EN_JUNCTION As Int16 = 0      ' Node types
    Public Const EN_RESERVOIR As Int16 = 1
    Public Const EN_TANK As Int16 = 2

    Public Const EN_CVPIPE As Int16 = 0        ' Link types
    Public Const EN_PIPE As Int16 = 1
    Public Const EN_PUMP As Int16 = 2
    Public Const EN_PRV As Int16 = 3
    Public Const EN_PSV As Int16 = 4
    Public Const EN_PBV As Int16 = 5
    Public Const EN_FCV As Int16 = 6
    Public Const EN_TCV As Int16 = 7
    Public Const EN_GPV As Int16 = 8

    Public Const EN_NONE As Int16 = 0          ' Quality analysis types
    Public Const EN_CHEM As Int16 = 1
    Public Const EN_AGE As Int16 = 2
    Public Const EN_TRACE As Int16 = 3

    Public Const EN_CONCEN As Int16 = 0        ' Source quality types
    Public Const EN_MASS As Int16 = 1
    Public Const EN_SETPOINT As Int16 = 2
    Public Const EN_FLOWPACED As Int16 = 3

    Public Const EN_CFS As Int16 = 0           ' Flow units types
    Public Const EN_GPM As Int16 = 1
    Public Const EN_MGD As Int16 = 2
    Public Const EN_IMGD As Int16 = 3
    Public Const EN_AFD As Int16 = 4
    Public Const EN_LPS As Int16 = 5
    Public Const EN_LPM As Int16 = 6
    Public Const EN_MLD As Int16 = 7
    Public Const EN_CMH As Int16 = 8
    Public Const EN_CMD As Int16 = 9

    Public Const EN_TRIALS As Int16 = 0       ' Misc. options
    Public Const EN_ACCURACY As Int16 = 1
    Public Const EN_TOLERANCE As Int16 = 2
    Public Const EN_EMITEXPON As Int16 = 3
    Public Const EN_DEMANDMULT As Int16 = 4

    Public Const EN_LOWLEVEL As Int16 = 0     ' Control types
    Public Const EN_HILEVEL As Int16 = 1
    Public Const EN_TIMER As Int16 = 2
    Public Const EN_TIMEOFDAY As Int16 = 3

    Public Const EN_AVERAGE As Int16 = 1      'Time statistic types
    Public Const EN_MINIMUM As Int16 = 2
    Public Const EN_MAXIMUM As Int16 = 3
    Public Const EN_RANGE As Int16 = 4

    Public Const EN_MIX1 As Int16 = 0         'Tank mixing models
    Public Const EN_MIX2 As Int16 = 1
    Public Const EN_FIFO As Int16 = 2
    Public Const EN_LIFO As Int16 = 3

    Public Const EN_NOSAVE As Int16 = 0       ' Save-results-to-file flag
    Public Const EN_SAVE As Int16 = 1
    Public Const EN_INITFLOW As Int16 = 10    ' Re-initialize flow flag

    'These are the external functions that comprise the DLL

    'Try replacing "Any" with "Int32()".
    'or Try replacing "Any" with "Object".

    Declare Function ENepanet Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal F1 As String, ByVal F2 As String, ByVal F3 As String, ByVal F4 As Object) As Long
    Declare Function ENopen Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal F1 As String, ByVal F2 As String, ByVal F3 As String) As Long
    Declare Function ENsaveinpfile Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal F As String) As Long
    Declare Function ENclose Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" () As Long

    Declare Function ENsolveH Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" () As Long
    Declare Function ENsaveH Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" () As Long
    Declare Function ENopenH Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" () As Long
    Declare Function ENinitH Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal SaveFlag As Long) As Long
    Declare Function ENrunH Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal T As Long) As Long
    Declare Function ENnextH Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal Tstep As Long) As Long
    Declare Function ENcloseH Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" () As Long
    Declare Function ENsavehydfile Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal F As String) As Long
    Declare Function ENusehydfile Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal F As String) As Long

    Declare Function ENsolveQ Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" () As Long
    Declare Function ENopenQ Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" () As Long
    Declare Function ENinitQ Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal SaveFlag As Long) As Long
    Declare Function ENrunQ Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal T As Long) As Long
    Declare Function ENnextQ Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal Tstep As Long) As Long
    Declare Function ENstepQ Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal Tleft As Long) As Long
    Declare Function ENcloseQ Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" () As Long

    Declare Function ENwriteline Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal S As String) As Long
    Declare Function ENreport Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" () As Long
    Declare Function ENresetreport Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" () As Long
    Declare Function ENsetreport Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal S As String) As Long

    Declare Function ENgetcontrol Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal Cindex As Long, ByVal Ctype_ As Long, ByVal Lindex As Long, ByVal Setting As Single, ByVal Nindex As Long, ByVal Level As Single) As Long
    Declare Function ENgetcount Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal Code As Long, ByVal Value As Long) As Long
    Declare Function ENgetoption Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal Code As Long, ByVal Value As Single) As Long
    Declare Function ENgettimeparam Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal Code As Long, ByVal Value As Long) As Long
    Declare Function ENgetflowunits Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal Code As Long) As Long
    Declare Function ENgetpatternindex Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal ID As String, ByVal Index As Long) As Long
    Declare Function ENgetpatternid Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal Index As Long, ByVal ID As String) As Long
    Declare Function ENgetpatternlen Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal Index As Long, ByVal L As Long) As Long
    Declare Function ENgetpatternvalue Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal Index As Long, ByVal Period As Long, ByVal Value As Single) As Long
    Declare Function ENgetqualtype Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal QualCode As Long, ByVal TraceNode As Long) As Long
    Declare Function ENgeterror Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal ErrCode As Long, ByVal ErrMsg As String, ByVal N As Long)

    Declare Function ENgetnodeindex Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal ID As String, ByVal Index As Long) As Long
    Declare Function ENgetnodeid Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal Index As Long, ByVal ID As String) As Long
    Declare Function ENgetnodetype Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal Index As Long, ByVal Code As Long) As Long
    Declare Function ENgetnodevalue Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal Index As Long, ByVal Code As Long, ByVal Value As Single) As Long

    Declare Function ENgetlinkindex Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal ID As String, ByVal Index As Long) As Long
    Declare Function ENgetlinkid Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal Index As Long, ByVal ID As String) As Long
    Declare Function ENgetlinktype Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal Index As Long, ByVal Code As Long) As Long
    Declare Function ENgetlinknodes Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal Index As Long, ByVal Node1 As Long, ByVal Node2 As Long) As Long
    Declare Function ENgetlinkvalue Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal Index As Long, ByVal Code As Long, ByVal Value As Single) As Long

    Declare Function ENgetversion Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal Value As Long) As Long

    Declare Function ENsetcontrol Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal Cindex As Long, ByVal Ctype_ As Long, ByVal Lindex As Long, ByVal Setting As Single, ByVal Nindex As Long, ByVal Level As Single) As Long
    Declare Function ENsetnodevalue Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal Index As Long, ByVal Code As Long, ByVal Value As Single) As Long
    Declare Function ENsetlinkvalue Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal Index As Long, ByVal Code As Long, ByVal Value As Single) As Long
    Declare Function ENaddpattern Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal ID As String) As Long
    Declare Function ENsetpattern Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal Index As Long, ByVal F As Object, ByVal N As Long) As Long
    Declare Function ENsetpatternvalue Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal Index As Long, ByVal Period As Long, ByVal Value As Single) As Long
    Declare Function ENsettimeparam Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal Code As Long, ByVal Value As Long) As Long
    Declare Function ENsetoption Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal Code As Long, ByVal Value As Single) As Long
    Declare Function ENsetstatusreport Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal Code As Long) As Long
    Declare Function ENsetqualtype Lib "C:\Users\user\Desktop\Desktop\epanet_toolkit_example_vb\epanet_toolkit_example_vb\bin\Debug\epanet2.dll" (ByVal QualCode As Long, ByVal ChemName As String, ByVal ChemUnits As String, ByVal TraceNode As String) As Long

End Module
