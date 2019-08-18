Imports CustomFormControls      'The programmer must also add a reference to CustomFormControls.dll in this project. You can find a copy of the dll in the frameworks installation directory.
Imports Complex_Number_Class    'The programmer must also add a reference to Complex_Number_Class in this project. You can find a copy of the dll in the frameworks installation directory.
Imports System.Xml
Imports System.IO
Imports System.Runtime.Serialization.Formatters.Binary

''' <summary>
''' Provides the caller with a list of the user-defined contents of the DLL.
''' </summary>
''' <remarks>These short descriptions will show up in various pull-down menus the user sees.
''' They are also used to tell the caller what functions are available in the DLL.</remarks>
Public Class UserDLLList
    ''' <summary>
    ''' Provide the caller with a list of the user-defined post processors in the DLL.
    ''' </summary>
    ''' <value></value>
    ''' <returns>A list of the descriptions that will appear in the pull-down menu.</returns>
    ''' <remarks>The names should be UserDLL_Posti with i=1,2,...</remarks>
    Public ReadOnly Property PostList() As String()
        Get
            Dim myPostList(1) As String
            myPostList(0) = "Dummy Post Processor"
            myPostList(1) = "X-Parameter Extractor"
            Return myPostList
        End Get
    End Property
    ''' <summary>
    ''' Provide the caller with a list of the user-defined post processors in the DLL.
    ''' </summary>
    ''' <value></value>
    ''' <returns>A list of the descriptions that will appear in the pull-down menu.</returns>
    ''' <remarks>The names should be UserDLLi with i=1,2,...</remarks>
    Public ReadOnly Property ModelList() As String()
        Get
            Dim myModelList(0) As String
            myModelList(0) = "Dummy User Model"
            Return myModelList
        End Get
    End Property
    ''' <summary>
    ''' Provide the caller with a list of the user-defined calibration engines in the DLL.
    ''' </summary>
    ''' <value></value>
    ''' <returns>A list of the descriptions that will appear in the pull-down menu.</returns>
    ''' <remarks>The names should be UserDLL_CalEngi with i=1,2,...</remarks>
    Public ReadOnly Property CalibrationEngineList() As String()
        Get
            Dim myCalibrationEngineList(0) As String
            myCalibrationEngineList(0) = "Dummy User SOLT Calibration Engine"
            Return myCalibrationEngineList
        End Get
    End Property
End Class

'User-defined Post-Processor example

''' <summary>
''' Post processor dummy
''' </summary>
''' <remarks></remarks>
Public Class UserDLL_Post1

    Private myPullDownSelection1 As Integer = -1
    Private myPullDownSelection2 As Integer = -1
    Private myPullDownSelection3 As Integer = -1
    Private myTextBoxContents As String = ""

    Private myNameList(1) As String   'The list of mechanism (model parameter) names for this model.

    'The user needs to intialize all of the values below.
    ''' <summary>
    ''' Set up the NameList for the model.
    ''' </summary>
    ''' <remarks>Use getNameList, getDescription, getTitle after initializeing.</remarks>
    Public Sub New()

        'Set up the list of input parameters for this post-processor.
        'This list will be displayed for the user and identify the role of each input parameter.
        myNameList(0) = "2x2 scattering parameters (.meas)"    'Dummy input scattering parameters.
        myNameList(1) = "Additional real gain factor (.parameter)"    'Dummy gain factor.

    End Sub
    ''' <summary>
    ''' Determines whether or not the multiple measurements specified in getMultipleMeasurements
    ''' are all gotten at once, or whether we loop through them
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>The default is false</remarks>
    Public ReadOnly Property getAllMultipleMeasurementsAtOnce As Boolean
        Get
            Return False
        End Get
    End Property
    ''' <summary>
    ''' Determines how the default frequency list is generated. 
    ''' </summary>
    ''' <value>Multiple input list = -2, Run>Set frequencies pull-down menu item = -1, list of input parameters = index of input parameter</value>
    ''' <returns></returns>
    ''' <remarks>You should add a description of how this is set to the overall description of this module. The user can override this default by checking the Run>Set frequencies pull-down menu item.</remarks>
    Public ReadOnly Property setFrequencyList() As Integer
        Get
            Return 0
        End Get
    End Property
    ''' <summary>
    ''' Determines how the default time list is generated. 
    ''' </summary>
    ''' <value>Multiple input list = -2, Run>Set times pull-down menu item = -1, list of input parameters = index of input parameter</value>
    ''' <returns></returns>
    ''' <remarks>You should add a description of how this is set to the overall description of this module if applicable. The user can override this default by checking the Run>Set times pull-down menu item.</remarks>
    Public ReadOnly Property setTimeList() As Integer
        Get
            Return -1
        End Get
    End Property
    ''' <summary>
    ''' A descriptive title for the text input required.
    ''' </summary>
    ''' <value></value>
    ''' <returns>A short string with a descriptive title for the text input required.</returns>
    ''' <remarks>Set to "" if you don't need user input from the text box.</remarks>
    Public ReadOnly Property getTextBoxDescription() As String
        Get
            Return ""
        End Get
    End Property
    ''' <summary>
    ''' Instruct the post-processor manager to run the post-processor multiple times on this input data
    ''' </summary>
    ''' <value>A string with a descriptive title of the multiple measurements.</value>
    ''' <returns></returns>
    ''' <remarks>Setting this blank disables the ability to treat multiple measurements.</remarks>
    Public ReadOnly Property getMultipleMeasurements() As String
        Get
            Return ""
        End Get
    End Property
    ''' <summary>
    ''' The contents of the text box entered by the user.
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property setTextBoxContents() As String
        Set(ByVal value As String)
            myTextBoxContents = value
        End Set
    End Property
    ''' <summary>
    ''' Tells the caller that the namelist must be reset when the pull-down selection is changed.
    ''' </summary>
    ''' <value></value>
    ''' <returns>True to force a reset of the namelist when the pull-down selection is changed.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getResetOnPullDownSelect() As Boolean()
        Get
            Dim ResetOnPullDownSelect(2) As Boolean
            ResetOnPullDownSelect(0) = False
            ResetOnPullDownSelect(1) = False
            ResetOnPullDownSelect(2) = False
            Return ResetOnPullDownSelect
        End Get
    End Property
    ''' <summary>
    ''' Set up the first pull-down list on the front panel.
    ''' </summary>
    ''' <value></value>
    ''' <returns>The selections in the pull-down list</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPullDownList1() As String()
        Get
            Dim PullDownList() As String = Nothing
            Return PullDownList
        End Get
    End Property
    ''' <summary>
    ''' Set up the second pull-down list on the front panel.
    ''' </summary>
    ''' <value></value>
    ''' <returns>The selections in the pull-down list</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPullDownList2() As String()
        Get
            Dim PullDownList() As String = Nothing  'No second pulldown list please
            Return PullDownList
        End Get
    End Property
    ''' <summary>
    ''' Set up the third pull-down list on the front panel.
    ''' </summary>
    ''' <value></value>
    ''' <returns>The selections in the pull-down list</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPullDownList3() As String()
        Get
            Dim PullDownList() As String = Nothing  'No third pulldown list please
            Return PullDownList
        End Get
    End Property
    ''' <summary>
    ''' The caller will set these before calling getRealMatrix or this instance will not know what the user selected.
    ''' Also, the caller 
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property setPullDownSelections() As Integer()
        Set(ByVal value As Integer())
            myPullDownSelection1 = value(0)
            myPullDownSelection2 = value(1)
            myPullDownSelection3 = value(2)
        End Set
    End Property
    ''' <summary>
    ''' Description for the model.
    ''' </summary>
    ''' <value>Model description</value>
    ''' <returns></returns>
    ''' <remarks>This appears below the list of model parameters on the front page of the model menu.
    ''' Please follow this example, starting with title, then a brief description, and then your author information.</remarks>
    Public ReadOnly Property getDescription() As String()
        Get
            'Setup a description for the model here. This will appear on the form under the model setup.
            Dim myDescription(2) As String
            myDescription(0) = "Dummy Post Processor. Returns real gain factor * S21 of the S-parameter input. The default frequency list is determined from the first input to the post processor."
            myDescription(2) = "Authored by Dylan Williams."
            Return myDescription
        End Get
    End Property
    ''' <summary>
    ''' Get the extension characterizing the result.
    ''' </summary>
    ''' <value></value>
    ''' <returns>The extension</returns>
    ''' <remarks>The options are .complex, .s1p, .s2p, and .s4p.
    ''' These will be bound up in a .meas object where the user can plot them, etc.
    ''' The types .complex and .s1p are the same, as they both have a single complex number at each frequency.
    ''' Real results should be saved as .complex with a zero imaginary part.</remarks>
    Public ReadOnly Property getResultExtension() As String
        Get
            Return ".complex"
        End Get
    End Property
    ''' <summary>
    ''' Get the NameList for this model
    ''' </summary>
    ''' <value></value>
    ''' <returns>The names for the mechanisms.</returns>
    ''' <remarks>This does not need to be customized by the programmer.</remarks>
    Public ReadOnly Property getNameList() As String()
        Get
            Return myNameList
        End Get
    End Property
    ''' <summary>
    ''' Select the file extensions that can be dropped as arguments into this Post Processor
    ''' </summary>
    ''' <value></value>
    ''' <returns>A list of the file extentions</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property SetFileExtensions() As String()
        Get

            Dim FileExtensions(45) As String : FileExtensions(0) = ".parameter" : FileExtensions(1) = ".meas" : FileExtensions(2) = ".model" : FileExtensions(3) = ".cascade" : FileExtensions(4) = ".sumofparameters" : FileExtensions(5) = ".waveform"  'This is the list of aceptable file extensions for the control.
            FileExtensions(6) = ".meas_archive" : FileExtensions(7) = ".model_archive" : FileExtensions(8) = ".cascade_archive" : FileExtensions(9) = ".sumofparameters" : FileExtensions(10) = ".sumofparameters_archive" : FileExtensions(11) = ".waveform_archive" : FileExtensions(12) = ".variables" : FileExtensions(13) = ".variables_archive"
            For i1 As Integer = 1 To 16 'Add in the .SNP AND .wnp wave files.
                FileExtensions(14 + 2 * (i1 - 1)) = ".s" + i1.ToString + "p"
                FileExtensions(15 + 2 * (i1 - 1)) = ".w" + i1.ToString + "p"
            Next i1
            Return FileExtensions
        End Get
    End Property
    ''' <summary>
    ''' Title of the model
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Appears on the front page of the menu.
    ''' Please start the name with something identifying the source of the model, such as the company or instituition acronym.
    ''' These titles should be short.</remarks>
    Public ReadOnly Property getTitle() As String
        Get
            getTitle = "Dummy Post Processor"
        End Get
    End Property

    ''' <summary>
    ''' Determines where the conditions associate with the output come from. 
    ''' </summary>
    ''' <value>Multiple input list = -1 and list of input parameters = index of input parameter</value>
    ''' <returns></returns>
    ''' <remarks>The number of condtions here must agree with the number of conditions in the ConditionNameList.</remarks>
    Public ReadOnly Property setConditionLocations() As Integer()
        Get
            Dim myConditionLocations() As Integer = Nothing
            Return myConditionLocations
        End Get
    End Property
    ''' <summary>
    ''' Determines the names of the conditions. 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Set a name to a blank to let the post processor pass all conditions from the input item.</remarks>
    Public ReadOnly Property setConditionNameList() As String()
        Get
            Dim myConditionNameList() As String = Nothing
            Return myConditionNameList
        End Get
    End Property

    ''' <summary>
    ''' Scattering parameters of the model
    ''' </summary>
    ''' <param name="MechValues">The list of input objects to the post processor.</param>
    ''' <param name="MechanismList1">The list of mechanisms we need to calculate the output.</param>
    ''' <param name="MultipleModelInput">The multiple input object to the post processor.</param>
    ''' <returns>The scattering parameters of the model.</returns>
    ''' <remarks>The mechanism list also has the frequency list inside.</remarks>
    Public Function getRealMatrix(ByVal MechValues() As Object, ByVal MechanismList1 As MechanismList, ByVal MultipleModelInput() As Object) As Object()


        'Get the imput real matrices. Here we only have one, the input sparameters.
        Dim myRealMatrix0 As RealMatrix : myRealMatrix0 = MechValues(0).getSParams(MechanismList1)
        Dim myRealMatrix1 As RealMatrix : myRealMatrix1 = MechValues(1).getSParams(MechanismList1)

        'Check the input to see if it has the right number of columns.
        If myRealMatrix0.NCols <> 9 Then MechanismList1.ErrorReport = "The first input to Dummy Post Processor was not a 2x2 scattering-parameter matrix" : MechanismList1.FatalError = True
        If myRealMatrix1.NCols <> 3 Then MechanismList1.ErrorReport = "The second input to Dummy Post Processor was not a scalar gain factor" : MechanismList1.FatalError = True

        'Form the real matrix output
        Dim myRealMatrixResult As New RealMatrix(myRealMatrix0.NRows, 3)
        myRealMatrixResult.Vector(1) = myRealMatrix0.Vector(1)
        For i As Integer = 1 To myRealMatrix0.NRows
            myRealMatrixResult(i, 2) = myRealMatrix0(i, 4) * myRealMatrix1(i, 2)
            myRealMatrixResult(i, 3) = myRealMatrix0(i, 5) * myRealMatrix1(i, 2)
        Next

        'That's all folks!
        Return ToArray(myRealMatrixResult)

    End Function

    ''' <summary>
    ''' A useful function for taking a real matrix output and putting it into the array format needed for the MUF
    ''' </summary>
    ''' <param name="InputMatrix">The input real matrix</param>
    ''' <returns>A one-element array of real matrices that MUF expects from the prost processors</returns>
    ''' <remarks></remarks>
    Private Function ToArray(ByRef InputMatrix As Object) As Object()
        Dim OutputMatrix(0) As Object
        OutputMatrix(0) = InputMatrix
        Return OutputMatrix
    End Function

End Class

''' <summary>
''' Post processor for extracting X-Parameters from intermediate files using the PNA-X or locally.
''' You must reference the AgilentNVNA Type Library (COM). This is available from the PNA-X and the 
''' help file gives instructions in the programming section for how to obtain this, and run commands 
''' remotely via DCOM.
''' 
''' Original code written by Laurence Stant in 2017 for n3m-labs.
''' </summary>
''' <remarks></remarks>
<Serializable()> Public Class UserDLL_Post2
    Implements ICloneable

    Private myPullDownSelection1 As Integer = -1
    Private myPullDownSelection2 As Integer = -1
    Private myPullDownSelection3 As Integer = -1
    Private myTextBoxContents As String = ""
    Private myNVNA As AgilentNVNA.Application
    Private myPNAXAddress As String = ""
    Private myLocalPath As String = ""
    Private myPNAXPath As String = ""

    Public Function Clone() As Object Implements System.ICloneable.Clone
        Dim m As New MemoryStream()
        Dim f As New BinaryFormatter()
        f.Serialize(m, Me)
        m.Seek(0, SeekOrigin.Begin)
        Return f.Deserialize(m)
    End Function

    'The user needs to intialize all of the values below.
    ''' <summary>
    ''' Set up the NameList for the model.
    ''' </summary>
    ''' <remarks>Use getNameList, getDescription, getTitle after initializeing.</remarks>
    Public Sub New()

    End Sub
    ''' <summary>
    ''' Determines whether or not the multiple measurements specified in getMultipleMeasurements
    ''' are all gotten at once, or whether we loop through them
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>The default is false</remarks>
    Public ReadOnly Property getAllMultipleMeasurementsAtOnce As Boolean
        Get
            Return False
        End Get
    End Property
    ''' <summary>
    ''' Instruct the post-processor manager to run the post-processor multiple times on this input data
    ''' </summary>
    ''' <value>A string with a descriptive title of the multiple measurements.</value>
    ''' <returns></returns>
    ''' <remarks>Setting this blank disables the ability to treat multiple measurements.</remarks>
    Public ReadOnly Property getMultipleMeasurements() As String
        Get
            Return ""
        End Get
    End Property
    ''' <summary>
    ''' A descriptive title for the text input required.
    ''' </summary>
    ''' <value></value>
    ''' <returns>A short string with a descriptive title for the text input required.</returns>
    ''' <remarks>Set to "" if you don't need user input from the text box.</remarks>
    Public ReadOnly Property getTextBoxDescription() As String
        Get
            Return "Options (see below):"
        End Get
    End Property
    ''' <summary>
    ''' The contents of the text box entered by the user.
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property setTextBoxContents() As String
        Set(ByVal value As String)
            myTextBoxContents = value
        End Set
    End Property
    ''' <summary>
    ''' Tells the caller that the namelist must be reset when the pull-down selection is changed.
    ''' </summary>
    ''' <value></value>
    ''' <returns>True to force a reset of the namelist when the pull-down selection is changed.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getResetOnPullDownSelect() As Boolean()
        Get
            Dim ResetOnPullDownSelect(2) As Boolean
            ResetOnPullDownSelect(0) = False
            ResetOnPullDownSelect(1) = False
            ResetOnPullDownSelect(2) = False
            Return ResetOnPullDownSelect
        End Get
    End Property
    ''' <summary>
    ''' Get the NameList for this model
    ''' </summary>
    ''' <value></value>
    ''' <returns>The names for the mechanisms.</returns>
    ''' <remarks>This needs to be customized by the programmer.</remarks>
    Public ReadOnly Property getNameList() As String()
        Get

            Dim myNameList(0) As String   'The list of mechanism (model parameter) names for this model.
            myNameList(0) = "Calibrated power waves (.meas)"

            Return myNameList

        End Get
    End Property
    ''' <summary>
    ''' Select the file extensions that can be dropped as arguments into this Post Processor
    ''' </summary>
    ''' <value></value>
    ''' <returns>A list of the file extentions</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property SetFileExtensions() As String()
        Get
            Dim FileExtensions(1) As String : FileExtensions(0) = ".meas" : FileExtensions(1) = ".meas_archive"
            Return FileExtensions
        End Get
    End Property
    ''' <summary>
    ''' Determines how the default frequency list is generated. 
    ''' </summary>
    ''' <value>Multiple input list = -2, Run>Set frequencies pull-down menu item = -1, list of input parameters = index of input parameter</value>
    ''' <returns></returns>
    ''' <remarks>You should add a description of how this is set to the overall description of this module. The user can override this default by checking the Run>Set frequencies pull-down menu item.</remarks>
    Public ReadOnly Property setFrequencyList() As Integer
        Get
            Return 1
        End Get
    End Property
    ''' <summary>
    ''' Determines how the default time list is generated. 
    ''' </summary>
    ''' <value>Multiple input list = -2, Run>Set times pull-down menu item = -1, list of input parameters = index of input parameter</value>
    ''' <returns></returns>
    ''' <remarks>You should add a description of how this is set to the overall description of this module if applicable. The user can override this default by checking the Run>Set times pull-down menu item.</remarks>
    Public ReadOnly Property setTimeList() As Integer
        Get
            Return -1
        End Get
    End Property
    ''' <summary>
    ''' Set up the first pull-down list on the front panel.
    ''' </summary>
    ''' <value></value>
    ''' <returns>The selections in the pull-down list</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPullDownList1() As String()
        Get
            Dim PullDownList(2) As String
            PullDownList(0) = "Use PNA-X for extraction"
            PullDownList(1) = "Use MUF extraction with phase-normalized waves"
            PullDownList(2) = "Use MUF extraction with raw waves"
            Return PullDownList
        End Get
    End Property
    ''' <summary>
    ''' Set up the second pull-down list on the front panel.
    ''' </summary>
    ''' <value></value>
    ''' <returns>The selections in the pull-down list</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPullDownList2() As String()
        Get
            Dim PullDownList() As String = Nothing 'No second pulldown list please
            Return PullDownList
        End Get
    End Property
    ''' <summary>
    ''' Set up the third pull-down list on the front panel.
    ''' </summary>
    ''' <value></value>
    ''' <returns>The selections in the pull-down list</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPullDownList3() As String()
        Get
            Dim PullDownList() As String = Nothing  'No third pulldown list please
            Return PullDownList
        End Get
    End Property
    ''' <summary>
    ''' The caller will set these before calling getRealMatrix or this instance will not know what the user selected.
    ''' Also, the caller 
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property setPullDownSelections() As Integer()
        Set(ByVal value As Integer())
            myPullDownSelection1 = value(0)
            myPullDownSelection2 = value(1)
            myPullDownSelection3 = value(2)
        End Set
    End Property
    ''' <summary>
    ''' Description for the model.
    ''' </summary>
    ''' <value>Model description</value>
    ''' <returns></returns>
    ''' <remarks>This appears below the list of model parameters on the front page of the model menu.
    ''' Please follow this example, starting with title, then a brief description, and then your author information.</remarks>
    Public ReadOnly Property getDescription() As String()
        Get
            'Setup a description for the model here. This will appear on the form under the model setup.
            Dim myDescription(4) As String
            myDescription(0) = "This post processor extracts X-Parameters from MDIF files containing power wave measurements calibrated by the MUF."
            myDescription(1) = "Currently 2-port devices are supported. For remote PNA-X extraction a mapped network drive is required to transfer data."
            myDescription(2) = "Please specify the PNA-X network address, local path, and remote path, in the Options textbox above, separated by ';' ."
            myDescription(3) = "Local path and remote path must point to the same location when navigated on this PC and the PNA-X, respectively."
            myDescription(4) = "Written by Laurence Stant, n3m-labs 2017."
            Return myDescription
        End Get
    End Property
    ''' <summary>
    ''' Get the extension characterizing the result.
    ''' </summary>
    ''' <value></value>
    ''' <returns>The extension</returns>
    ''' <remarks>The options are .complex, .s1p, .s2p, and .s4p.
    ''' These will be bound up in a .meas object where the user can plot them, etc.
    ''' The types .complex and .s1p are the same, as they both have a single complex number at each frequency.
    ''' Real results should be saved as .complex with a zero imaginary part.</remarks>
    Public ReadOnly Property getResultExtension() As String
        Get
            Return ".xnp"
        End Get
    End Property
    ''' <summary>
    ''' Title of the model
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Appears on the front page of the menu.
    ''' Please start the name with something identifying the source of the model, such as the company or instituition acronym.
    ''' These titles should be short.</remarks>
    Public ReadOnly Property getTitle() As String
        Get
            getTitle = "Extract X-parameters from wave data in MDIF format."
        End Get
    End Property


    ''' <summary>
    ''' Determines where the conditions associate with the output come from. 
    ''' </summary>
    ''' <value>Multiple input list = -1 and list of input parameters = index of input parameter</value>
    ''' <returns></returns>
    ''' <remarks>The number of condtions here must agree with the number of conditions in the ConditionNameList.</remarks>
    Public ReadOnly Property setConditionLocations() As Integer()
        Get
            Dim myConditionLocations() As Integer = Nothing
            Return myConditionLocations
        End Get
    End Property
    ''' <summary>
    ''' Determines the names of the conditions. 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Set a name to a blank to let the post processor pass oll of the conditions of the item through the post process to the output.</remarks>
    Public ReadOnly Property setConditionNameList() As String()
        Get
            Dim myConditionNameList() As String = Nothing
            Return myConditionNameList
        End Get
    End Property

    ''' <summary>
    ''' Device model parameters determined by ADS
    ''' </summary>
    ''' <param name="MechValues">The list of input objects to the post processor.</param>
    ''' <param name="MechanismList1">The list of mechanisms we need to calculate the output.</param>
    ''' <param name="MultipleModelInput">The multiple input object to the post processor.</param>
    ''' <returns>The scattering parameters of the model.</returns>
    ''' <remarks>The mechanism list also has the frequency list inside.</remarks>
    Public Function getRealMatrix(ByVal MechValues() As Object, ByVal MechanismList1 As MechanismList, ByVal MultipleModelInput() As Object) As Object()

        ' Initialise the input MDIF file so we have something to return if things go wrong!
        Dim myMDIF As New MDIF()
        myMDIF = getMDIF(MechValues(0), MechanismList1)

        If MechanismList1.InitializeFlag And myPullDownSelection1 = 0 Then
            ' First run and we are using PNA-X, so set up communication

            If (myTextBoxContents.Trim() = "") Then
                ' Assume local extraction - MUF running on PNA-X. Trim guards against invisible spaces when 'empty'
                myPNAXAddress = "localhost"
                myLocalPath = IO.Path.GetTempPath
                myPNAXPath = myLocalPath
            Else
                ' Separate the text box arguments
                Dim myTextBoxParts = myTextBoxContents.Split(";")
                If (myTextBoxParts.Count = 3) Then
                    ' Assume remote PNA-X specified
                    myPNAXAddress = myTextBoxParts(0)
                    myLocalPath = myTextBoxParts(1)
                    myPNAXPath = myTextBoxParts(2)
                Else
                    ' Not sure what we have - error
                    MechanismList1.FatalError = True : MechanismList1.ErrorReport = "Error: Please supply zero or three arguments in the text box."
                    Return ToArray(myMDIF) : Exit Function
                End If
            End If


            ' Try and talk to the PNA-X
            Try
                PNAX_Initialize_XP_Extraction(myPNAXAddress)

            Catch ex As System.IO.FileNotFoundException ' Catches passed up to here as I need access to MechanismList1 for error handling
                MechanismList1.FatalError = True : MechanismList1.ErrorReport = "Error: Cannot connect to NVNA Application at specified address."
                Return ToArray(myMDIF) : Exit Function
            Catch ex As System.NotSupportedException
                MechanismList1.FatalError = True : MechanismList1.ErrorReport = "Error: X-Parameters not supported on specified PNA-X."
                Return ToArray(myMDIF) : Exit Function
            Catch
                MechanismList1.FatalError = True : MechanismList1.ErrorReport = "Error: Cannot start instrument driver. Check NVNAComServer is registered."
                Return ToArray(myMDIF) : Exit Function
            End Try

        End If

        ' Now lets extract X-parameters.

        Dim myRealMatrixResult As Object = Nothing 'Initialise result object

        If myPullDownSelection1 = 0 Then

            ' Perform PNA-X extraction
            Try
                myRealMatrixResult = PNAX_Extract_XPs(myMDIF, myPNAXAddress, myLocalPath, myPNAXPath)

            Catch ex As System.IO.InvalidDataException
                MechanismList1.FatalError = True : MechanismList1.ErrorReport = "PNA-X could not extract X-parameters! Check file contents."
                Return ToArray(myMDIF) : Exit Function

            Catch
                MechanismList1.FatalError = True : MechanismList1.ErrorReport = "Error: Cannot connect to PNA-X NVNA Application."
                Return ToArray(myMDIF) : Exit Function
            End Try

        ElseIf myPullDownSelection1 = 1 Then

            ' Perform MUF extraction with phase-normalized waves
            myRealMatrixResult = MUF_Extract_XPs(myMDIF, True)

        ElseIf myPullDownSelection1 = 2 Then

            ' Perform MUF extraction with raw waves
            myRealMatrixResult = MUF_Extract_XPs(myMDIF, False)

        End If

        ' Return the result
        Return ToArray(myRealMatrixResult)

    End Function

    Private Sub PNAX_Initialize_XP_Extraction(myPNAXAddress As String)

        Try
            myNVNA = CreateObject("AgilentNVNA.Application", myPNAXAddress)
            If IsNothing(myNVNA) Then
                Throw New System.IO.FileNotFoundException ' Most suitable?
            End If
            myNVNA.Preset()
            myNVNA.XparameterEnabled = True
            If myNVNA.XparameterEnabled = False Then
                Throw New System.NotSupportedException
            End If
        Catch
            Throw ' Pass up to caller
        End Try

    End Sub

    Private Function PNAX_Extract_XPs(myMDIF As MDIF, myPNAXAddress As String, myLocalPath As String, myPNAXPath As String) As Object

        ' Write MDIF file for PNA-X to access
        myMDIF.Write(IO.Path.Combine(myLocalPath, "dut.mdf"))

        Dim success As Boolean = myNVNA.GenerateXParamFromFiles(IO.Path.Combine(myPNAXPath, "dut.mdf"), IO.Path.Combine(myPNAXPath, "dut.xnp"), False)
        If Not success Then
            Throw New System.IO.InvalidDataException
        End If

        ' Read in result
        Dim myXNP As New MDIF
        myXNP.Read(IO.Path.Combine(myLocalPath, "dut.xnp"))

        Return myXNP

    End Function

    Private Function MUF_Extract_XPs(myMDIF As MDIF, normalize_phase As Boolean) As Object

        ' We can either sweep through AN_1_1 of each stimulus tone and write out each block at a time to an xnp file,
        ' or build a big array of values and then write them out altogether. We will use the latter for now.

        ' 1. Set blockVAR index to 0
        ' 1b. Get shape of ET states from first blockVARs
        ' 2. Increment blockVAR index,  if valid get values of indepVARs
        ' 3.    Get block indices of that set of indepVARs
        ' 4.    Build index of ET states
        ' 5.    Extract X-Parameters using this index
        ' 6. Loop
        ' 7. Write X-Parameters to file

        Dim blockVAR_index As Integer = 0 ' 1. Set blockVAR index to 0
        Dim current_block_VARs As HPList = Nothing
        Dim ET_vars As String() = {"ssport", "ssfreq", "ssphase"}

        ' 1b. Get shape of ET states from first blockVARs
        Dim ssports As Integer = 1
        Dim ssfreqs As Integer = 1
        Dim ssphases As Integer = 1
        Dim current_ssport As Double = 1
        Dim current_ssfreq As Double = 0
        Dim current_ssphase As Double = 1
        While True
            current_block_VARs = myMDIF.BlockVARs(blockVAR_index)
            For i As Integer = 0 To current_block_VARs.count - 1
                Dim name As String = current_block_VARs.GetHPName(i)
                Dim value As String = current_block_VARs.GetValueDouble(i)
                If ET_vars.Contains(name) Then
                    Select Case name
                        Case "ssport"
                            If value < current_ssport Then
                                Exit While
                            End If
                            If value > current_ssport Then
                                ssports = ssports + 1
                                current_ssport = value
                            End If
                        Case "ssfreq"
                            If value > current_ssfreq Then
                                ssfreqs = ssfreqs + 1
                                current_ssfreq = value
                            End If
                        Case "ssphase"
                            If value > current_ssphase Then
                                ssphases = ssphases + 1
                                current_ssphase = value
                            End If
                        Case Else

                    End Select

                End If
            Next
            blockVAR_index = blockVAR_index + 1
        End While

        Dim n_X_params As Integer = ((ssfreqs - 1) * ssports * 2 + 1) * (ssfreqs - 1) * ssports 'XFpk, XSpkql, XTpkql
        Dim X_params As New ComplexMatrix(myMDIF.BlockCount, n_X_params) ' We'll trim the rows later
        Dim X_param_block_indices(myMDIF.BlockCount) As Integer ' And these rows
        Dim X_param_index As Integer = 1
        blockVAR_index = 0

        While True
            ' 2. Increment blockVAR index,  if valid get values of indepVARs.
            If (blockVAR_index = myMDIF.BlockCount) Then
                ' We've got through all the stimulus conditions!
                Exit While
            Else

                current_block_VARs = myMDIF.BlockVARs(blockVAR_index)
                Dim indepVar_sweep_array(current_block_VARs.count - 4) As MDIF_Var_Sweep
                Dim j As Integer = 0
                For i As Integer = 0 To current_block_VARs.count - 1

                    Dim name As String = current_block_VARs.GetHPName(i)
                    Dim value As Double = current_block_VARs.GetValueDouble(i)

                    ' Unless it's the ET variables...
                    If ET_vars.Contains(name) Then
                        Continue For
                    End If

                    ' Add the indepVar to our sweep object array
                    indepVar_sweep_array(j) = New MDIF_Var_Sweep(name, value, value)
                    j += 1
                Next

                ' 3.    Get block indices of that set of indepVARs
                Dim ET_states As Integer() = myMDIF.GetBlockIndexFromVarRanges(indepVar_sweep_array)

                ' 4.    Build index of ET states
                Dim ET_index(ssports - 1, ssfreqs - 1, ssphases - 1) As Integer
                Dim index As Integer = 0
                For ssport As Integer = 0 To ssports - 1
                    For ssfreq As Integer = 0 To ssfreqs - 1
                        For ssphase As Integer = 0 To ssphases - 1
                            ET_index(ssport, ssfreq, ssphase) = ET_states(index)
                            index = index + 1
                        Next
                    Next
                Next

                ' 5.    Extract X-Parameters using this index

                ' Fill matrices

                Dim B_s(ssports - 1, ssfreqs - 1, ssphases - 1) As ComplexMatrix
                Dim A_s(ssports - 1, ssfreqs - 1, ssphases - 1) As ComplexMatrix

                For ssport As Integer = 0 To ssports - 1
                    For ssfreq As Integer = 0 To ssfreqs - 1
                        For ssphase As Integer = 0 To ssphases - 1
                            Dim block_index As Integer = ET_index(ssport, ssfreq, ssphase)
                            Dim this_block As RealMatrix = myMDIF.BlockMatrix(block_index).CreateRealMatrix
                            Dim A As New ComplexMatrix(ssfreqs - 1, ssports)
                            Dim B As New ComplexMatrix(ssfreqs - 1, ssports)
                            Dim P As Complex
                            P = toComplex(this_block.Rarray(0, 1), this_block.Rarray(0, 2))
                            P = P / Abs(P)
                            For port As Integer = 0 To ssports - 1
                                For freq As Integer = 0 To ssfreqs - 2
                                    ' this_block: freq, A1 real, A1 imag, B1 real, B1 imag, A2 real, A2 imag.
                                    ' Add one to complex matrix indices because they are 1 indexed for some reason...
                                    A(freq + 1, port + 1) = toComplex(this_block.Rarray(freq, port * 4 + 1), this_block.Rarray(freq, port * 4 + 2)) '* P ^ (freq + 1)
                                    B(freq + 1, port + 1) = toComplex(this_block.Rarray(freq, port * 4 + 3), this_block.Rarray(freq, port * 4 + 4)) '* P ^ (freq + 1)
                                    A(freq + 1, port + 1) = A(freq + 1, port + 1) + New Complex(1.0E-17 * (port + 1), 1.0E-17)
                                    B(freq + 1, port + 1) = B(freq + 1, port + 1) + New Complex(1.0E-17 * (port + 1), 1.0E-17)
                                Next
                            Next
                            A_s(ssport, ssfreq, ssphase) = A
                            B_s(ssport, ssfreq, ssphase) = B
                        Next
                    Next
                Next

                ' Next step
                Dim X_columns As Integer = (ssfreqs - 1) * ssports * 2 + 1 - 1 '-1 as we are fitting XSpk11 and XTpk11 together
                Dim X As New ComplexMatrix(ssports * ssfreqs * ssphases - ssphases, X_columns) '-1 as we don't include ET on A11
                Dim Y As New ComplexMatrix(ssports * ssfreqs * ssphases - ssphases)
                Dim ET_i As Integer
                Dim A0 As New Complex(0, 0)
                Dim A0s As New ComplexMatrix(ssfreqs - 1, ssports)
                Dim s As New ComplexMatrix(X_columns)

                ' Calculate A0
                Dim OPT_average_A0 As Boolean = False
                If OPT_average_A0 Then
                    For ET_port As Integer = 0 To ssports - 1
                        For ET_phase As Integer = 0 To ssphases - 1
                            A0 = A0 + A_s(ET_port, 0, ET_phase)(1, 1) / (ssports * ssphases)
                        Next
                    Next
                Else
                    A0 = A0 + A_s(0, 0, 0)(1, 1)
                End If

                For port As Integer = 0 To ssports - 1
                    For freq As Integer = 0 To ssfreqs - 2
                        ET_i = 1
                        For ssport As Integer = 0 To ssports - 1
                            For ssfreq As Integer = 0 To ssfreqs - 1
                                If ssport = 0 And ssfreq = 1 Then Continue For
                                For ssphase As Integer = 0 To ssphases - 1
                                    Y(ET_i) = B_s(ssport, ssfreq, ssphase)(freq + 1, port + 1)
                                    X(ET_i, 1) = toComplex(1, 0)
                                    For a_port As Integer = 0 To ssports - 1
                                        For a_freq As Integer = 0 To ssfreqs - 2
                                            If a_port = 0 And a_freq = 0 Then
                                                X(ET_i, (a_port * (ssfreqs - 1) + a_freq) + 2) = (A_s(ssport, ssfreq, ssphase)(a_freq + 1, a_port + 1) - A0) + New Complex(1.0E-17, 1.0E-17)
                                            Else
                                                X(ET_i, (a_port * (ssfreqs - 1) + a_freq) + 2) = A_s(ssport, ssfreq, ssphase)(a_freq + 1, a_port + 1)
                                                X(ET_i, (a_port * (ssfreqs - 1) + a_freq) + (ssports * ssfreqs - 1)) = Conj(A_s(ssport, ssfreq, ssphase)(a_freq + 1, a_port + 1))
                                            End If
                                        Next
                                    Next
                                    ET_i += 1
                                Next
                            Next
                        Next

                        ' LSE
                        s = ((ConjTranspose(X) * X) ^ -1) * (ConjTranspose(X) * Y)

                        'XF
                        'X_params(X_param_index, port * (ssfreqs - 1) + freq + 1) = s(1)
                        Dim XF As New Complex(0, 0)
                        XF = B_s(0, 0, 0)(freq + 1, port + 1)
                        For ET_port As Integer = 0 To ssports - 1
                            'For ET_phase As Integer = 0 To ssphases - 1
                            '    XF = XF + B_s(ET_port, 0, ET_phase)(freq + 1, port + 1) / (ssports * ssphases)
                            'Next
                            For ET_freq As Integer = 0 To ssfreqs - 2
                                If ET_port = 0 And ET_freq = 0 Then
                                    'XF = XF - s(2 + (ET_port * (ssfreqs - 1) + ET_freq)) * (A_s(0, 0, 0)(freq + 1, port + 1) - A0)
                                Else
                                    XF = XF - s(2 + (ET_port * (ssfreqs - 1) + ET_freq)) * A_s(0, 0, 0)(freq + 1, port + 1)
                                    XF = XF - s(1 + (ET_port * (ssfreqs - 1) + ET_freq) + (ssports * (ssfreqs - 1))) * Conj(A_s(0, 0, 0)(freq + 1, port + 1))
                                End If
                            Next
                        Next

                        X_params(X_param_index, port * (ssfreqs - 1) + freq + 1) = XF
                        For q As Integer = 0 To ssports - 1
                            For l As Integer = 0 To ssfreqs - 2
                                If q = 0 And l = 0 Then
                                    'XSpk11 = XS + XT
                                    X_params(X_param_index, ssports * (ssfreqs - 1) + port * (ssfreqs - 1) * ssports * (ssfreqs - 1) + freq * ssports * (ssfreqs - 1) + q * (ssfreqs - 1) + l + 1) = s(2 + (q * (ssfreqs - 1) + l))
                                    'XTpk11 = 0
                                    X_params(X_param_index, ssports * (ssfreqs - 1) + port * (ssfreqs - 1) ^ 2 * ssports + freq * ssports * (ssfreqs - 1) + q * (ssfreqs - 1) + l + 1 + (ssfreqs - 1) ^ 2 * ssports ^ 2) = New Complex(0, 0)
                                Else
                                    'XS
                                    X_params(X_param_index, ssports * (ssfreqs - 1) + port * (ssfreqs - 1) * ssports * (ssfreqs - 1) + freq * ssports * (ssfreqs - 1) + q * (ssfreqs - 1) + l + 1) = s(2 + (q * (ssfreqs - 1) + l))
                                    'XT
                                    X_params(X_param_index, ssports * (ssfreqs - 1) + port * (ssfreqs - 1) ^ 2 * ssports + freq * ssports * (ssfreqs - 1) + q * (ssfreqs - 1) + l + 1 + (ssfreqs - 1) ^ 2 * ssports ^ 2) = s(1 + (q * (ssfreqs - 1) + l) + (ssports * (ssfreqs - 1)))
                                End If
                            Next
                        Next

                    Next
                Next

                X_param_block_indices(X_param_index - 1) = blockVAR_index
                X_param_index = X_param_index + 1
                blockVAR_index = ET_states(ET_states.Length - 1) + 1 'Jump next loop index to next set of indepVARs

            End If

        End While

        ' Can't redim preserve X_params so there are lots of empty rows below X_param_index...
        ReDim Preserve X_param_block_indices(X_param_index - 2)
        Dim last_X_param_index As Integer = X_param_index - 1

        ' Now write out our X-Parameter file as an MDIF!

        Dim xnp As New MDIF
        xnp.DataType = ".xnp"

        'Dim X_param_block As New ComplexMatrix(0, X_params.NCols)
        Dim X_param_block As New RealMatrix(0, 1 + 2 * X_params.NCols)

        Dim current_BlockVARs As HPList = myMDIF.BlockVARs(X_param_block_indices(0))
        Dim current_AN_1_1 As Double = current_BlockVARs.GetValueDouble("AN_1_1")
        Dim AN_1_1_array As Double()
        Dim current_AN_1_1_sqW As Double
        Dim new_AN_1_1 As Double
        X_param_index = 1 ' Use 1 because of ComplexMatrix indexing...
        Dim block_start As Integer = 1
        Dim block_end As Integer

        'Now we add the "header" blocks. **I think we want to do this at the end** to get info on indices and NumFundFreqs etc.

        'XParamAttributes
        Dim XPA_Index As Double = 0 'Not used so far, but should update for each block I think
        Dim XPA_Version As Double = 2 'Used by ADS. Subtle changes.
        Dim XPA_NumPorts As Double = 2 'Should be automated.
        Dim XPA_NumFundFreqs As Double = 1 'Should also be automated.
        Dim XPA_Names As String = "% Index(integer) Version(real) NumPorts(integer) NumFundFreqs(integer)"
        Dim XParamAttributes_Matrix As New HPMatrix("") 'Used because the below is a nice oneliner and we can't set a row easily
        XParamAttributes_Matrix.AddNamesToMatrix(XPA_Names)
        XParamAttributes_Matrix.StoreMatrixLines(CStr(XPA_Index) + " " + CStr(XPA_Version) + " " + CStr(XPA_NumPorts) + " " + CStr(XPA_NumFundFreqs))
        xnp.AddBlock("XParamAttributes", New HPList(""), New HPList(""), XParamAttributes_Matrix.CreateRealMatrix)
        'Have to add Names afterwards as we can't pass a HPMatrix to AddBlock, only a RealMatrix
        xnp.BlockMatrix(0).AddNamesToMatrix(XPA_Names)

        'XParamPortData
        Dim XPPD_RefZ0 As Integer = 50
        Dim XParamPortData_Matrix As New HPMatrix("")
        Dim XPPD_Names As String = "% PortNumber(integer) RefZ0(complex) PortName(string)"
        XParamPortData_Matrix.AddNamesToMatrix(XPPD_Names)
        XParamPortData_Matrix.StoreMatrixLines("1 50 0 ""Port 1""")
        XParamPortData_Matrix.StoreMatrixLines("2 50 0 ""Port 2""")
        xnp.AddBlock("XParamPortData", New HPList(""), New HPList(""), XParamPortData_Matrix.CreateRealMatrix)
        xnp.BlockMatrix(1) = XParamPortData_Matrix 'Must overwrite as can't pass string using RealMatrix through AddBlock

        While True
            ' Assume innermost indepVar is AN_1_1 and sweep that as our row
            ' Also assume it always increases
            current_BlockVARs = myMDIF.BlockVARs(X_param_block_indices(X_param_index - 1))
            new_AN_1_1 = current_BlockVARs.GetValueDouble("AN_1_1")
            If (new_AN_1_1 < current_AN_1_1 Or X_param_index = last_X_param_index) Then
                If (new_AN_1_1 < current_AN_1_1) Then
                    ' X_param_index is on the start of the next block
                    block_end = X_param_index - 1
                Else
                    ' X_param_index is on the end of the block - we hit the end of the file
                    block_end = X_param_index
                    Array.Resize(AN_1_1_array, X_param_index - block_start + 1)
                    AN_1_1_array(X_param_index - block_start) = Math.Sqrt(10 ^ ((new_AN_1_1 - 30) / 10))
                End If

                'If current_BlockVARs.GetHPName(current_BlockVARs.count - 1) = "CVindx" Then
                '    current_BlockVARs.DeleteVariable(current_BlockVARs.count - 6) '4 if MCindx and CV indx not present...
                'Else
                '    current_BlockVARs.DeleteVariable(current_BlockVARs.count - 4) '4 if MCindx and CV indx not present...
                'End If

                Dim X_param_block_VARs As New HPList("")
                Dim end_of_kept_VARs As Integer = 1
                If current_BlockVARs.GetHPName(current_BlockVARs.count - 1) = "CVindx" Then
                    end_of_kept_VARs = current_BlockVARs.count - 7
                Else
                    end_of_kept_VARs = current_BlockVARs.count - 5
                End If
                For i As Integer = 0 To end_of_kept_VARs
                    X_param_block_VARs.AddToVariables("VAR " + current_BlockVARs.GetHPName(i) + "(" + current_BlockVARs.GetHPType(i) + ") = " + current_BlockVARs.GetValueString(i))
                Next

                X_param_block.ReDimension(block_end - block_start + 1, 1 + 2 * X_params.NCols)
                For i As Integer = 1 To block_end - block_start + 1
                    ' Not sure where we get AN_1_1 in the xnp file from yet - let's take it from the AN_1_1 in the MDIF file for now
                    current_AN_1_1_sqW = AN_1_1_array(i - 1)
                    X_param_block(i, 1) = current_AN_1_1_sqW
                    For j As Integer = 2 To X_params.NCols + 1
                        X_param_block(i, 2 * (j - 1)) = X_params(block_start + i - 1, j - 1).Re
                        X_param_block(i, 2 * (j - 1) + 1) = X_params(block_start + i - 1, j - 1).Im
                    Next
                Next
                xnp.AddBlock("XParamData", X_param_block_VARs, New HPList(""), X_param_block)
                xnp.BlockMatrix(xnp.BlockCount - 1).AddNamesToMatrix("% AN_1_1(real) FB_1_1(complex) FB_1_2(complex) FB_1_3(complex) FB_2_1(complex) FB_2_2(complex) FB_2_3(complex) S_1_1_1_1(complex) S_1_1_1_2(complex) S_1_1_1_3(complex) S_1_1_2_1(complex) S_1_1_2_2(complex) S_1_1_2_3(complex) S_1_2_1_1(complex) S_1_2_1_2(complex) S_1_2_1_3(complex) S_1_2_2_1(complex) S_1_2_2_2(complex) S_1_2_2_3(complex) S_1_3_1_1(complex) S_1_3_1_2(complex) S_1_3_1_3(complex) S_1_3_2_1(complex) S_1_3_2_2(complex) S_1_3_2_3(complex) S_2_1_1_1(complex) S_2_1_1_2(complex) S_2_1_1_3(complex) S_2_1_2_1(complex) S_2_1_2_2(complex) S_2_1_2_3(complex) S_2_2_1_1(complex) S_2_2_1_2(complex) S_2_2_1_3(complex) S_2_2_2_1(complex) S_2_2_2_2(complex) S_2_2_2_3(complex) S_2_3_1_1(complex) S_2_3_1_2(complex) S_2_3_1_3(complex) S_2_3_2_1(complex) S_2_3_2_2(complex) S_2_3_2_3(complex) T_1_1_1_1(complex) T_1_1_1_2(complex) T_1_1_1_3(complex) T_1_1_2_1(complex) T_1_1_2_2(complex) T_1_1_2_3(complex) T_1_2_1_1(complex) T_1_2_1_2(complex) T_1_2_1_3(complex) T_1_2_2_1(complex) T_1_2_2_2(complex) T_1_2_2_3(complex) T_1_3_1_1(complex) T_1_3_1_2(complex) T_1_3_1_3(complex) T_1_3_2_1(complex) T_1_3_2_2(complex) T_1_3_2_3(complex) T_2_1_1_1(complex) T_2_1_1_2(complex) T_2_1_1_3(complex) T_2_1_2_1(complex) T_2_1_2_2(complex) T_2_1_2_3(complex) T_2_2_1_1(complex) T_2_2_1_2(complex) T_2_2_1_3(complex) T_2_2_2_1(complex) T_2_2_2_2(complex) T_2_2_2_3(complex) T_2_3_1_1(complex) T_2_3_1_2(complex) T_2_3_1_3(complex) T_2_3_2_1(complex) T_2_3_2_2(complex) T_2_3_2_3(complex)")

                block_start = X_param_index
            End If

            Array.Resize(AN_1_1_array, X_param_index - block_start + 1)
            AN_1_1_array(X_param_index - block_start) = Math.Sqrt(10 ^ ((new_AN_1_1 - 30) / 10))

            X_param_index = X_param_index + 1
            current_AN_1_1 = new_AN_1_1

            If X_param_index > last_X_param_index Then Exit While

        End While

        Return xnp

    End Function

    ''' <summary>
    ''' A useful function for taking a real matrix output and putting it into the array format needed for the MUF
    ''' </summary>
    ''' <param name="InputMatrix">The input real matrix</param>
    ''' <returns>A one-element array of real matrices that MUF expects from the prost processors</returns>
    ''' <remarks></remarks>
    Private Function ToArray(ByRef InputMatrix As Object) As Object()
        Dim OutputMatrix(0) As Object
        OutputMatrix(0) = InputMatrix
        Return OutputMatrix
    End Function

    Public Function getMDIF(measSupport As MeasurementSupport, ByVal MechanismList1 As MechanismList) As MDIF

        Dim NominalMDIF As MDIF = Nothing
        Dim NominalMDIFFull As Boolean = False

        'Get the root element and node of the menu.
        Dim RootElement As XmlElement = measSupport.myDoc.DocumentElement
        Dim RootNode As XmlNode = measSupport.myDoc.DocumentElement

        'Get the SParameter value. There should only be one.
        Dim elem As XmlElement = RootNode.SelectSingleNode("/CorrectedMeasurement/Controls/MeasSParams/Item[@Index='0']/SubItem[@Index='1']")
        If elem Is Nothing Then elem = RootNode.SelectSingleNode("/Measurement/Controls/MeasSParams/Item[@Index='0']/SubItem[@Index='1']")

        'The nominal value is what we have found so far.
        Dim myValue As XmlElement = elem    'This is the default node for the scattering parameters.
        Dim NominalValue As Boolean = True  'Keep track of whther this is the nominal value.

        'Get the mechanism name. There should only be one.
        elem = RootNode.SelectSingleNode("/CorrectedMeasurement/Controls/MeasurementName")
        If elem Is Nothing Then elem = RootNode.SelectSingleNode("/Measurement/Controls/MeasurementName")
        Dim ValueString As String = elem.GetAttribute("ControlText")
        Dim MeasurementName As String = ValueString

        'Now see if we should select a pertrubed scattering parameter instead.
        If Not MechanismList1 Is Nothing Then

            If MechanismList1.IsMonteCarloSimulation Then

                'A Monte-Carlo, use the subindex
                'Get the index in the mechanism list and the subindex if any.
                Dim myMechanismIndex As Integer = MechanismList1.Index(MeasurementName)
                If myMechanismIndex >= 0 Then

                    Dim myMechanismSubIndex As Integer = MechanismList1.SubIndex(myMechanismIndex)

                    'See if we need to look for a pertrubed value
                    If MechanismList1.IsPerturbed(myMechanismIndex) And Not MechanismList1.InitializeFlag Then

                        'Get the correct file from the list of MonteCarlo s-parameters.
                        elem = RootNode.SelectSingleNode("/CorrectedMeasurement/Controls/MonteCarloPerturbedSParams")
                        If elem Is Nothing Then elem = RootNode.SelectSingleNode("/Measurement/Controls/MonteCarloPerturbedSParams")
                        ValueString = elem.GetAttribute("Count")
                        Dim MeasurementCount As Integer = Val(ValueString)

                        If MeasurementCount > 0 Then

                            'Recycle the Monte-Carlo simulations if we don't have enough of them.
                            While myMechanismSubIndex >= MeasurementCount
                                myMechanismSubIndex = myMechanismSubIndex - MeasurementCount
                            End While

                            'Get the value corresponding to the Monte-Carlo simulation
                            If myMechanismSubIndex >= 0 And myMechanismSubIndex < MeasurementCount Then

                                'Seems that we have found a perturbed value
                                myValue = elem.SelectSingleNode("Item[@Index='" + myMechanismSubIndex.ToString + "']/SubItem[@Index='1']")
                                NominalValue = False

                            End If

                        End If

                    End If

                End If

            Else

                'A covariance result. Cycle through the list of measurements we have.
                'If we find one with IsPerturbed=True, use it.

                'Get the number of Covariance s-parameters.
                elem = RootNode.SelectSingleNode("/CorrectedMeasurement/Controls/PerturbedSParams")
                If elem Is Nothing Then elem = RootNode.SelectSingleNode("/Measurement/Controls/PerturbedSParams")
                ValueString = elem.GetAttribute("Count")
                Dim MeasurementCount As Integer = Val(ValueString)

                'Cycle through the list looking for a perturbed value.
                If MeasurementCount > 0 Then
                    For i As Integer = 0 To MeasurementCount - 1

                        'Get the mechanism name for each item on the list
                        elem = RootNode.SelectSingleNode("/CorrectedMeasurement/Controls/PerturbedSParams/Item[@Index='" + i.ToString + "']/SubItem[@Index='2']")
                        If elem Is Nothing Then elem = RootNode.SelectSingleNode("/Measurement/Controls/PerturbedSParams/Item[@Index='" + i.ToString + "']/SubItem[@Index='2']")
                        ValueString = elem.GetAttribute("Text")
                        'See if the mechanism name is pertrubed
                        Dim myIndex As Integer = MechanismList1.Index(ValueString)
                        If myIndex >= 0 Then
                            If MechanismList1.IsPerturbed(myIndex) Then
                                'We found a perturbed mechanism. Update myValue
                                myValue = RootNode.SelectSingleNode("/CorrectedMeasurement/Controls/PerturbedSParams/Item[@Index='" + i.ToString + "']/SubItem[@Index='1']")
                                If myValue Is Nothing Then myValue = RootNode.SelectSingleNode("/Measurement/Controls/PerturbedSParams/Item[@Index='" + i.ToString + "']/SubItem[@Index='1']")
                                NominalValue = False
                            End If
                        End If

                    Next i
                End If

            End If
        End If

        'Actually get the real matrix we need to return
        If Not NominalMDIFFull Or Not NominalValue Or MechanismList1.InitializeFlag Or MechanismList1.PreSolutionFlag Then    'Read the value from disk

            'Create an XML document corresponding to myValue and pass it to FileMeasurementSupport to get the scattering parameters.
            'Get the xml node below controlsubitem
            Dim MechanismXmlElement As XmlElement = myValue.FirstChild

            'Make a document SubDoc loaded from disk for this mechanism???
            Dim SubDoc As New XmlDocument
            Dim dec As XmlDeclaration = SubDoc.CreateXmlDeclaration("1.0", Nothing, Nothing)
            SubDoc.AppendChild(dec)

            Dim xmlChildNode As XmlNode = SubDoc.ImportNode(MechanismXmlElement, True)
            SubDoc.AppendChild(xmlChildNode)
            'SubDoc.Save("c:\temp.txt")

            'Use FileMeasurementSupport to generate the scattering parameters we need.
            Dim newFileMeasurementSupport As New FileMeasurementSupport(SubDoc)
            Dim temp As MDIF = newFileMeasurementSupport.getMDIF()    'Get the data in FileMeasurementSupport rather than just a copy to save time.
            Dim FileMDIF As New MDIF : FileMDIF = temp.Clone          'This copy will stay around for longer, as MeasurementSupport endures.

            'Let's save our nominal value if we have one for later use.
            If NominalValue Then
                NominalMDIF = New MDIF
                NominalMDIF = temp.Clone
                NominalMDIFFull = True
            End If

            Return FileMDIF

        Else    'Special case. We already saved a nominal value and we are not do a presolution or initializing. Let's save some time and just return it to the caller.

            Dim FileMDIF As New MDIF : FileMDIF = NominalMDIF.Clone
            Return FileMDIF

        End If

    End Function

End Class

'User-defined Model example

''' <summary>
''' Dummy
''' </summary>
''' <remarks></remarks>
Public Class UserDLL1

    Private myNameList() As String = Nothing  'The list of mechanism (model parameter) names for this model.

    'The user needs to intialize all of the values below.
    ''' <summary>
    ''' Set up the NameList for the model.
    ''' </summary>
    ''' <remarks>Use getNameList, getDescription, getTitle after initializeing.</remarks>
    Public Sub New()

        'No namelist for this one.

    End Sub
    ''' <summary>
    ''' Description for the model.
    ''' </summary>
    ''' <value>Model description</value>
    ''' <returns></returns>
    ''' <remarks>This appears below the list of model parameters on the front page of the model menu.
    ''' Please follow this example, starting with title, then a brief description, and then your author information.</remarks>
    Public ReadOnly Property getDescription() As String()
        Get
            'Setup a description for the model here. This will appear on the form under the model setup.
            Dim myDescription(2) As String
            myDescription(0) = "Software templates are available for writing new models. Writing new models is not difficult, and all the software needed is free. Start by creating a user model for yourself. Once your code is tested, send it to Dylan for possible inclusion in the next release."
            myDescription(2) = "Look in the help files for more information."
            Return myDescription
        End Get
    End Property
    ''' <summary>
    ''' Get the NameList for this model
    ''' </summary>
    ''' <value></value>
    ''' <returns>The names for the mechanisms.</returns>
    ''' <remarks>This is the only part of the class that does not need to be customized by the programmer.</remarks>
    Public ReadOnly Property getNameList() As String()
        Get
            Return myNameList
        End Get
    End Property
    ''' <summary>
    ''' Title of the model
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Appears on the front page of the menu.
    ''' Please start the name with something identifying the source of the model, such as the company or instituition acronym.
    ''' These titles should be short.</remarks>
    Public ReadOnly Property getTitle() As String
        Get
            getTitle = "Dummy"
        End Get
    End Property
    ''' <summary>
    ''' Gets the extension we want to use with this model.
    ''' Choises are ".model" or ".operator"
    ''' </summary>
    ''' <value></value>
    ''' <returns>The extension</returns>
    ''' <remarks>Use .model for simple models that return 2x2 scattering paraemters.
    ''' Use .operator for models that perform a transformation on 2x2 s-parameters.
    ''' The only difference is the file extention. But it will help the user figure out what is going on.</remarks>
    Public ReadOnly Property getExtension() As String
        Get
            Return ".model"
        End Get
    End Property
    ''' <summary>
    ''' Scattering parameters of the model
    ''' </summary>
    ''' <param name="MechValues">The values of the mechanisms (i.e. mocel parameters) in the mechanism list requested by the model.</param>
    ''' <param name="MechanismList1">The list of mechanisms (i.e. mocel parameters) we need to calculate the scattering parameters.</param>
    ''' <returns>The scattering parameters of the model.</returns>
    ''' <remarks>The mechanism list also has the frequency list inside.</remarks>
    Public Function getSParams(ByVal MechValues() As Double, ByVal MechanismList1 As MechanismList) As RealMatrix
        'Calculate the models scattering parameters
        Dim myOne As New RealMatrix(MechanismList1.FrequencyList.NRows, 9, ".s2p") : myOne.Vector(1) = MechanismList1.FrequencyList : myOne.InitializeAsSParams()  'Scattering parameters.
        Return getSParams(MechValues, MechanismList1, myOne)
    End Function
    ''' <summary>
    ''' Scattering parameters of the model
    ''' </summary>
    ''' <param name="MechValues">The values of the mechanisms (i.e. mocel parameters) in the mechanism list requested by the model.</param>
    ''' <param name="MechanismList1">The list of mechanisms (i.e. mocel parameters) we need to calculate the scattering parameters.</param>
    ''' <param name="SParamsIn">The input scattering parameters to which these scattering parameters are cascaded.</param>
    ''' <returns>The scattering parameters of the model.</returns>
    ''' <remarks>The mechanism list also has the frequency list inside.</remarks>
    Public Function getSParams(ByVal MechValues() As Double, ByVal MechanismList1 As MechanismList, ByVal SParamsIn As RealMatrix) As RealMatrix

        'Create a scattering-parameter instance.
        Dim mySParams As New RealMatrix(MechanismList1.FrequencyList.NRows, 9, ".s2p") : mySParams.Vector(1) = MechanismList1.FrequencyList

        'A useful matrix for holding single-frequency s-parameters.
        Dim S As New ComplexMatrix(2, 2)

        'The mechanism values

        'Step through the frequencies and generate the scattering parameters of the junction.
        For k As Integer = 1 To mySParams.NRows

            'Set up the frequencies and other parameters.
            Dim FGHz As Double = mySParams(k, 1)

            'Call the actual routine.
            S(1, 1) = toComplex(0.0, 0.0) : S(2, 2) = S(1, 1)
            S(2, 1) = toComplex(1.0, 0.0) : S(1, 2) = S(2, 1)

            'Store the result
            mySParams.SMatrix(k) = S

        Next k

        'Cascade the input scattering parameters if they are supplied.
        mySParams = mySParams.CascadeSParameters(SParamsIn)

        'Return the values to the caller. That's all folks!
        Return mySParams

    End Function

End Class

'User-defined Calibration Engine example
Public Class UserDLL_CalEng1
    Inherits CalibrationEngine1

    Public Sub New()

    End Sub

    ''' <summary>
    ''' A description of the calibration engine used by the caller to fill out the final VNA pane
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared ReadOnly Property Description() As String()
        Get
            Dim myDescription(4) As String
            myDescription(0) = "The SOLT calibration algorithm is very simple and extremely robust. You need at least three one-port standards on each port, and a thru connection. The three one-port standards must be well separated on the Smith Chart at all frequencies, or you must add additional one-port standards until at least three of the one-port standards are well separated at every frequency. This algorithm uses a least-squares fit on each port using the one-port standards, followed by a determination of the thru parameters based on the thru connection.  The one-port standards must come in pairs, with one one-port standard on each port."
            myDescription(1) = ""
            myDescription(2) = "If you use only thru standards, the thru will be used to set the transmission tracking, and must be fully defined. If you use a reciprocal thru, or a combination of reciprocal thrus and regular thrus, the transmission tracking terms will be determined from the termination standards. The thrus and reciprocal thrus will only be used to adjust the ratio of the forward and reverse transmission tracking terms."
            myDescription(3) = ""
            myDescription(4) = "Written by Dylan Williams"
            Return myDescription
        End Get
    End Property
    ''' <summary>
    ''' A list of the standard types supported by this calibration engine.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>These standard types appear in the pull-down menu the user sees on the main panel.
    ''' Later they must be used to identify the standard types for the calibration engine.
    ''' These can be supplied in any order, so it is best to put the most commonly used on top of the list.</remarks>
    Public Shared ReadOnly Property StandardTypes() As String()
        Get
            Dim myStdTypes(4) As String
            myStdTypes(0) = "Thru"
            myStdTypes(1) = "Termination (S21=S12=0)"
            myStdTypes(2) = "Isolation standard"
            myStdTypes(3) = "Switch terms"
            myStdTypes(4) = "Reciprocal thru"
            Return myStdTypes
        End Get
    End Property
    ''' <summary>
    ''' A list of the length requirements for the standard types supported by this calibration engine.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>This helps the program better indicate to the user when lengths are required </remarks>
    Public Shared ReadOnly Property LengthNeeded() As Boolean()
        Get
            Dim myLengthNeeded(4) As Boolean
            myLengthNeeded(0) = False ' "Thru"
            myLengthNeeded(1) = False ' "Termination"
            myLengthNeeded(2) = False ' "Isolation standard"
            myLengthNeeded(3) = False ' "Switch terms"
            myLengthNeeded(4) = False ' "Reciprocal thru"
            Return myLengthNeeded
        End Get
    End Property
    ''' <summary>
    ''' Tells the caller if this program needs to create its own menu
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared ReadOnly Property MenuNeeded() As Boolean
        Get
            Return False
        End Get
    End Property
    ''' <summary>
    ''' Tells the caller if this program will be forcing values
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared ReadOnly Property WillForceValues() As Boolean
        Get
            Return False
        End Get
    End Property
    ''' <summary>
    ''' Set true for multiport algorithms that understand how to solve a multiport problem.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>If false, input to algoithm will be .s2p or .w2p and output will be .s4p.
    ''' If true, input to algorithm will be .snp or .wnp and output will be .s2np.</remarks>
    Public Shared ReadOnly Property AlgorithmIsMultiport() As Boolean
        Get
            Return False
        End Get
    End Property
    ''' <summary>
    ''' Set true for algorithms that use wave representations instead of scattering-parameter representations for measurements.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>If true, the algorithm will get two .wnp files in for each standard.
    ''' If false, the algorithm will get one .snp file in for each standard.</remarks>
    Public Shared ReadOnly Property AlgorithIsWaveRepresentation() As Boolean
        Get
            Return False
        End Get
    End Property
    ''' <summary>
    ''' Set true for one-port algorithms.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Only makes sense if this is not a multiport algorithm.
    ''' If AlgorithmIsMultiport = False and AlgorithmIsOnePort = False, then the algorithm is a standard 2-port algorithm.</remarks>
    Public Shared ReadOnly Property AlgorithmIsOneport() As Boolean
        Get
            Return False
        End Get
    End Property
    ''' <summary>
    ''' This routine is called to close the calibration engine down.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CloseEngine()
    End Sub

    ''' <summary>
    ''' Run SOLT calibration
    ''' </summary>
    ''' <param name="SimulationCount">SimulationCount is passed from the caller, and is used to number the results</param>
    ''' <param name="myInitializeMenuFile">This is the path containing the initialization menu. Not relevant for this type of calibration.</param>
    ''' <param name="myRootPath">This is the path containing the menu and data files.</param>
    ''' <param name="myFinalResultsPath">This is where the final results get stored.</param>
    ''' <param name="MechanismList1">The mechanism list</param>
    ''' <remarks></remarks>
    Public Sub RunCalibration(ByVal SimulationCount As Integer, ByVal myInitializeMenuFile As String, ByVal myRootPath As String, ByVal myFinalResultsPath As String, ByRef MechanismList1 As MechanismList)

        'In this program we will set up a StatistiCAL like solution file as well.
        'This was done to allow some code reuse, but is optional.
        'Most programs will probably create the results needed to fill up myNewMechanismList without going through this step.
        Dim FileNameSolution As String = myFinalResultsPath + "\Solution_" + SimulationCount.ToString + ".txt"

        'Define useful variables
        Dim NTerm As Integer = 0, NThru As Integer = 0
        Dim SParams As New RealMatrix(MechanismList1.FrequencyList.NRows, 9, ".s2p") : SParams.Vector(1) = MechanismList1.FrequencyList : SParams.InitializeAsSParams()
        Dim SParamsIso As New RealMatrix(MechanismList1.FrequencyList.NRows, 9, ".switch")
        Dim SParamsSwitch As New RealMatrix(MechanismList1.FrequencyList.NRows, 9, ".iso")
        Dim Solution As New RealMatrix(MechanismList1.FrequencyList.NRows, 41, ".solution")

        'Count the terminations and get the switch and isolation terms.
        Dim NTerms As Integer = 0, NThrus As Integer = 0, ReciprocalThru As Boolean = False
        For i As Integer = 1 To myCalibrationModels.Count
            If Not myIgnoreCalibrationModels(i - 1) Then
                Select Case myCalibrationStdTypes(i - 1)
                    Case "Termination (S21=S12=0)"
                        NTerms = NTerms + 1
                    Case "Isolation standard"
                        SParamsIso = myCalibrationMeasurements(i - 1).getSParams(MechanismList1)
                        MechanismList1.IsolationTerms = SParamsIso
                    Case "Switch terms"
                        SParamsSwitch = myCalibrationMeasurements(i - 1).getSParams(MechanismList1)
                        MechanismList1.SwitchTerms = SParamsSwitch
                    Case "Thru"
                        NThrus = NThrus + 1
                    Case "Reciprocal thru"
                        NThrus = NThrus + 1
                        ReciprocalThru = True
                End Select
            End If
        Next i

        'Create matrices to hold the definitions and measurements for ports 1 and 2.
        Dim ThruMeas(NThrus) As RealMatrix
        Dim ThruDef(NThrus) As RealMatrix
        Dim TerminationDef11 As New ComplexMatrix(MechanismList1.FrequencyList.NRows, NTerms)
        Dim TerminationMeas11 As New ComplexMatrix(MechanismList1.FrequencyList.NRows, NTerms)
        Dim TerminationDef22 As New ComplexMatrix(MechanismList1.FrequencyList.NRows, NTerms)
        Dim TerminationMeas22 As New ComplexMatrix(MechanismList1.FrequencyList.NRows, NTerms)
        Dim Base1SwitchCorrectedMeas As New RealMatrix(MechanismList1.FrequencyList.NRows, 9)

        'Cycle throught the calibration standards and create matrices with the termination definitions and measurements.
        Solution.Vector(1) = MechanismList1.FrequencyList
        For i As Integer = 1 To myCalibrationModels.Count
            If Not myIgnoreCalibrationModels(i - 1) Then
                Select Case myCalibrationStdTypes(i - 1)
                    Case "Termination (S21=S12=0)" 'This is a "Load"
                        NTerm = NTerm + 1
                        SParams = EquivalentDefinition(myCalibrationModels(i - 1), MechanismList1)
                        TerminationDef11.Vector(NTerm) = SParams.Vector(2) + toComplex(0.0, 1.0) * SParams.Vector(3)
                        TerminationDef22.Vector(NTerm) = SParams.Vector(8) + toComplex(0.0, 1.0) * SParams.Vector(9)
                        SParams = myCalibrationMeasurements(i - 1).getSParams(MechanismList1)
                        Base1SwitchCorrectedMeas = SParams.SwitchTermCorrection(SParams, MechanismList1)
                        TerminationMeas11.Vector(NTerm) = Base1SwitchCorrectedMeas.Vector(2) + toComplex(0.0, 1.0) * Base1SwitchCorrectedMeas.Vector(3)
                        TerminationMeas22.Vector(NTerm) = Base1SwitchCorrectedMeas.Vector(8) + toComplex(0.0, 1.0) * Base1SwitchCorrectedMeas.Vector(9)
                    Case "Thru", "Reciprocal thru"  'This is the StatistiCAL Reciprocal Adapter or bend (S22, S22 known, S12 unknown)
                        NThru = NThru + 1
                        ThruMeas(NThru) = New RealMatrix(MechanismList1.FrequencyList.NRows, 9)
                        ThruDef(NThru) = New RealMatrix(MechanismList1.FrequencyList.NRows, 9)
                        Dim Temp As New RealMatrix(MechanismList1.FrequencyList.NRows, 9)
                        SParams = myCalibrationMeasurements(i - 1).getSParams(MechanismList1)
                        Temp = SParams.SwitchTermCorrection(SParams, MechanismList1)
                        ThruMeas(NThru).Fill(Temp) '= SParams.SwitchTermCorrection(SParams.Base1StoredSParams, MechanismList1)
                        SParams = EquivalentDefinition(myCalibrationModels(i - 1), MechanismList1)
                        Temp = SParams
                        ThruDef(NThru).Fill(Temp) ' = SParams.Base1StoredSParams
                End Select
            End If
        Next i

        'Swap rows and columns to put the matrices in a more convinient form.
        Dim TerminationDef11t As New ComplexMatrix(NTerms, MechanismList1.FrequencyList.NRows)
        Dim TerminationMeas11t As New ComplexMatrix(NTerms, MechanismList1.FrequencyList.NRows)
        Dim TerminationDef22t As New ComplexMatrix(NTerms, MechanismList1.FrequencyList.NRows)
        Dim TerminationMeas22t As New ComplexMatrix(NTerms, MechanismList1.FrequencyList.NRows)
        TerminationDef11t = Transpose(TerminationDef11)
        TerminationMeas11t = Transpose(TerminationMeas11)
        TerminationDef22t = Transpose(TerminationDef22)
        TerminationMeas22t = Transpose(TerminationMeas22)

        'Get the guess if we have one. In this case, we only use it to check that we are consistent with our signs of S1_21
        Dim PreSolution As New RealMatrix(MechanismList1.FrequencyList.NRows, 33, ".solution")
        Dim PreSolutionNx16 As New ComplexMatrix(MechanismList1.FrequencyList.NRows, 16)
        PreSolution = MechanismList1.PreSolution
        For i As Integer = 1 To MechanismList1.FrequencyList.NRows
            For k1 As Integer = 1 To 16
                PreSolutionNx16(i, k1) = toComplex(MechanismList1.PreSolution(i, 2 + 2 * (k1 - 1)), MechanismList1.PreSolution(i, 3 + 2 * (k1 - 1))) 'S11
            Next k1
        Next i

        'Run the SOLT engine and solve for the solution vector.
        For i As Integer = 1 To MechanismList1.FrequencyList.NRows

            'Move the current results into a 4x4 complex matrix
            Dim Pre4x4 As New ComplexMatrix(4, 4)
            For i1 As Integer = 1 To 4
                For j1 As Integer = 1 To 4
                    Pre4x4(i1, j1) = PreSolutionNx16(i, 1 + 4 * (i1 - 1) + (j1 - 1))
                Next j1
            Next i1
            Dim PreS1_21 As New Complex : PreS1_21 = Pre4x4(1, 3)

            'Solve the least-squares problem.
            Dim SP1_11 As New Complex(0.0, 0.0), SP1_22 As New Complex(0.0, 0.0), SP1_T As New Complex(0.0, 0.0), SP1_det As New Complex(0.0, 0.0)
            Call Unterm(SP1_11, SP1_22, SP1_det, TerminationMeas11t.Vector(i), TerminationDef11t.Vector(i))
            SP1_T = Complex_Number_Class.Sqrt(SP1_11 * SP1_22 - SP1_det)
            If (SP1_T / PreS1_21).Re < 0.0 Then SP1_T = -SP1_T 'Swap signs of S1_21 if it does not agree with sign selected in presolution.
            Dim SP2_11 As New Complex(0.0, 0.0), SP2_22 As New Complex(0.0, 0.0), SP2_T As New Complex(0.0, 0.0), SP2_det As New Complex(0.0, 0.0)
            Call Unterm(SP2_22, SP2_11, SP2_det, TerminationMeas22t.Vector(i), TerminationDef22t.Vector(i))
            SP2_T = Complex_Number_Class.Sqrt(SP2_11 * SP2_22 - SP2_det)

            'Get the transmission terms
            Dim k As New Complex(1.0, 0.0), k_est(NThrus) As Complex, SP2_T_est(NThrus) As Complex : NThru = 0
            If ReciprocalThru Then  'We are only setting k. S2_21*S2_12 determined from transmission terms

                'Determine k from the ratio of the thru's forward to reverse transmission coef for each thru
                For istd As Integer = 1 To myCalibrationModels.Count
                    If Not myIgnoreCalibrationModels(istd - 1) And (myCalibrationStdTypes(istd - 1) = "Thru" Or myCalibrationStdTypes(istd - 1) = "Reciprocal thru") Then

                        'Setup
                        NThru = NThru + 1
                        Dim S21_meas As New Complex(ThruMeas(NThru)(i, 4), ThruMeas(NThru)(i, 5)), S12_meas As New Complex(ThruMeas(NThru)(i, 6), ThruMeas(NThru)(i, 7))
                        k_est(NThru) = Complex_Number_Class.Sqrt(S21_meas / S12_meas)
                        Dim S21 As New Complex(0.0, 0.0), S12 As New Complex(0.0, 0.0)
                        Dim S21_def As New Complex(ThruDef(NThru)(i, 4), ThruDef(NThru)(i, 5)), S12_def As New Complex(ThruDef(NThru)(i, 6), ThruDef(NThru)(i, 7))
                        Dim S11_def As New Complex(ThruDef(NThru)(i, 2), ThruDef(NThru)(i, 3)), S22_def As New Complex(ThruDef(NThru)(i, 8), ThruDef(NThru)(i, 9))

                        'Calcuate the denominator (see Terminology for High-Speed Sampling-Oscilloscope Calibration, eq. 9)
                        Dim den As New Complex : den = (1.0 - SP1_22 * S11_def - SP2_11 * S22_def - SP1_22 * SP2_11 * (S12_def * S21_def - S11_def * S22_def))
                        S21 = SP1_T * (SP2_T * k_est(NThru)) / den
                        S12 = SP1_T * (SP2_T / k_est(NThru)) / den

                        'Reconcile the signs of sp2_21 with those of the thru
                        Dim ratio As New Complex(0.0, 0.0)
                        ratio = S21 / S21_meas
                        If ratio.Re < 0.0 Then k_est(NThru) = -k_est(NThru)

                        ''Check the reconciliation
                        'S21 = SP1_T * (SP2_T * k(NThru)) / den
                        'S12 = SP1_T * (SP2_T / k(NThru)) / den

                    End If
                Next istd

                'Now average the k's
                k = toComplex(0.0, 0.0)
                For ik As Integer = 1 To NThrus
                    k = k + k_est(ik)
                Next ik
                k = (1.0 / CDbl(NThrus)) * k

            Else    'Determine S2_21 and S2_12 from the thru.

                'Determine the transmission terms from the thru measurements
                For istd As Integer = 1 To myCalibrationModels.Count
                    If Not myIgnoreCalibrationModels(istd - 1) And (myCalibrationStdTypes(istd - 1) = "Thru" Or myCalibrationStdTypes(istd - 1) = "Reciprocal thru") Then

                        'Setup
                        NThru = NThru + 1
                        Dim S21_meas As New Complex(ThruMeas(NThru)(i, 4), ThruMeas(NThru)(i, 5)), S12_meas As New Complex(ThruMeas(NThru)(i, 6), ThruMeas(NThru)(i, 7))
                        Dim S11_meas As New Complex(ThruMeas(NThru)(i, 2), ThruMeas(NThru)(i, 3)), S22_meas As New Complex(ThruMeas(NThru)(i, 8), ThruMeas(NThru)(i, 9))
                        Dim S21 As New Complex(0.0, 0.0), S12 As New Complex(0.0, 0.0)
                        Dim S21_def As New Complex(ThruDef(NThru)(i, 4), ThruDef(NThru)(i, 5)), S12_def As New Complex(ThruDef(NThru)(i, 6), ThruDef(NThru)(i, 7))
                        Dim S11_def As New Complex(ThruDef(NThru)(i, 2), ThruDef(NThru)(i, 3)), S22_def As New Complex(ThruDef(NThru)(i, 8), ThruDef(NThru)(i, 9))

                        'Note that these equations "push" all of the changes into SP2_T, not SP1_T. It would be nice if the correction could be made symmetric.
                        'Calcuate the denominator (see Terminology for High-Speed Sampling-Oscilloscope Calibration, eq. 9)
                        Dim den As New Complex : den = (1.0 - SP1_22 * S11_def - SP2_11 * S22_def - SP1_22 * SP2_11 * (S12_def * S21_def - S11_def * S22_def))
                        S21 = S21_meas / (SP1_T * S21_def / den)
                        S12 = S12_meas / (SP1_T * S12_def / den)

                        'Convert to the StatistiCAL format
                        k_est(NThru) = Complex_Number_Class.Sqrt(S21 / S12)
                        SP2_T_est(NThru) = Complex_Number_Class.Sqrt(S21 * S12)

                        'Resolve any sign difficulties
                        If NThru = 1 Then
                            'The first time we check for internally-consistent sign chioces
                            Dim Ratio As New Complex(1.0, 0.0)
                            Ratio = S21 / (SP2_T_est(NThru) * k_est(NThru))
                            If Ratio.Re < 0.0 Then k_est(NThru) = -k_est(NThru)
                        Else
                            'After that, we have to check to see that we are consistent with the sign choices we made the first time
                            Dim Ratio As New Complex(1.0, 0.0)
                            Ratio = SP2_T_est(NThru) / SP2_T_est(1)
                            If Ratio.Re < 0.0 Then SP2_T_est(NThru) = -SP2_T_est(NThru)
                            Ratio = k_est(NThru) / k_est(1)
                            If Ratio.Re < 0.0 Then k_est(NThru) = -k_est(NThru)
                        End If

                        ''Check the reconciliation
                        'S21 = SP1_T * (SP2_T * k(NThru)) / den
                        'S12 = SP1_T * (SP2_T / k(NThru)) / den

                    End If
                Next istd

                'Now average the k's and SP2_T's
                k = toComplex(0.0, 0.0) : SP2_T = toComplex(0.0, 0.0)
                For ik As Integer = 1 To NThrus
                    k = k + k_est(ik)
                    SP2_T = SP2_T + SP2_T_est(ik)
                Next ik
                k = (1.0 / CDbl(NThrus)) * k
                SP2_T = (1.0 / CDbl(NThrus)) * SP2_T

            End If

            'Stuff the results into the solution file
            'This is not necessary, but rather done for convience.
            Solution(i, 2) = SP1_11.Re : Solution(i, 3) = SP1_11.Im
            Solution(i, 4) = SP1_22.Re : Solution(i, 5) = SP1_22.Im
            Solution(i, 6) = SP1_T.Re : Solution(i, 7) = SP1_T.Im
            Solution(i, 8) = SP2_11.Re : Solution(i, 9) = SP2_11.Im
            Solution(i, 10) = SP2_22.Re : Solution(i, 11) = SP2_22.Im
            Solution(i, 12) = SP2_T.Re : Solution(i, 13) = SP2_T.Im
            Solution(i, 14) = k.Re : Solution(i, 15) = k.Im 'k = sqrt(S2_21/S2_12)
            Solution(i, 16) = 1.0 'Dummy eps eff

        Next i

        'Save the solution file to disk with the name Solution.txt
        Complex_Number_Class.Write(Solution, FileNameSolution)

        'Put 4x4 scattering parameter matrix that calibrates the VNA or LSNA into the new solution
        'For convienience, we use the routine already for generating this from a StatistiCAL solution vector.
        Dim NewSolution As New RealMatrix(MechanismList1.FrequencyList.NRows, 33, ".s4p")
        NewSolution.Read4x4SEfromSolutionVector(FileNameSolution)
        MechanismList1.PreSolution = NewSolution

    End Sub

    ''' <summary>
    ''' Subroutine Unterm performs a one-port SOLT
    ''' </summary>
    ''' <param name="S11">Error box S11</param>
    ''' <param name="S22">Error box S22</param>
    ''' <param name="Ds">Error box determinant.</param>
    ''' <param name="G_in">The measured scattering parameters.</param>
    ''' <param name="G_L">The definitions of the standards.</param>
    ''' <remarks>See Notebook 5, page 39. This formulation avoids treating G_L=0 as a special case.</remarks>
    Private Sub Unterm(ByRef S11 As Complex, ByRef S22 As Complex, ByRef Ds As Complex, ByVal G_in As ComplexMatrix, ByVal G_L As ComplexMatrix)

        'Fill the alpha and beta matrix.
        Dim Beta As New ComplexMatrix(G_in.NRows, 3), Alpha As New ComplexMatrix(G_in.NRows), Prod As New ComplexMatrix(G_in.NRows)
        For i As Integer = 1 To G_in.NRows
            Alpha(i) = -G_in(i)
            Beta(i, 1) = toComplex(-1.0, 0.0)
            Beta(i, 2) = -G_in(i) * G_L(i)
            Beta(i, 3) = G_L(i)
        Next i

        'Solve for p=(Beta^T *Beta)^-1 * Beta^T * Alpha where ^T is complex transpose.
        Dim p As New ComplexMatrix(3), BetaT As New ComplexMatrix(3, G_in.NRows)
        BetaT = ConjTranspose(Beta)
        p = ((BetaT * Beta) ^ -1) * (BetaT * Alpha)

        'Separate out the results
        S11 = p(1) : S22 = p(2) : Ds = p(3)

    End Sub

    ''' <summary>
    ''' The user wishes to initialize the menu. Set up a menu in InitalizeDirectory
    ''' </summary>
    ''' <param name="InitalizeDirectory">The directory where the menu gets stored</param>
    ''' <param name="InitializeMenu">= 1: Create menu from scratch and let user modify it.
    ''' = 2: Menu should already exist. Let the user modify it.</param>
    ''' <param name="MechanismList1">The Mechanism List</param>
    ''' <remarks>The Initialized menu must be saved to the file InitalizeDirectory + "\Initialize_Menu.txt" 
    ''' Simple calibration engines won't need menus to set options, and so won't need to do anything here.</remarks>
    Public Sub InitializeMenu(ByVal InitalizeDirectory As String, ByVal InitializeMenu As Integer, ByVal MechanismList1 As MechanismList)

    End Sub

End Class

'This class is required for compatibility and should not be changed.

Public Class CalibrationEngine1

    'The information we will use to do the calibration.
    Protected Friend myNumberOfPorts As Integer              'True if the algorithm can measure multiports
    Protected Friend myPort1Connection() As Integer          'The port that the corresponding standard is connected to
    Protected Friend myPort2Connection() As Integer          'The port that the corresponding standard is connected to
    Private myCurrentPort1Connection As Integer = -1         'The current port connections. Set to -1 to ensure that these are set somewhere before using them.
    Private myCurrentPort2Connection As Integer = -1
    Protected Friend myPortModels(7) As Object               'The port parameters TestPort1, TestPort2, ...
    Protected Friend myDUTPortModels(7) As Object            'The DUT port parameters DUTPort1, DUTPort2, ...
    Protected Friend myCalibrationModels() As Object         'The models serving as the calibration definitions
    Protected Friend myIgnoreCalibrationModels() As Boolean  'True = don't use this standard in the simulation
    Protected Friend myCalibrationModelNames() As String     'The models serving as the calibration definitions
    Protected Friend myCalibrationStdTypes() As String       'The strings defining the type of calibration standards
    Protected Friend myCalibrationLengths() As Object        'The mechanisms that define the length (if applicable) of the calibration standard
    Protected Friend myCalibrationMeasurements() As Object   'The object that returns the raw measurement of this standard
    Protected Friend myDUTMeasurements() As Object           'The objects that return the raw measurements of the DUTs
    Protected Friend myDUTMeasurementNames() As String       'The strings with the DUT names. These are used to create the subdirectories for the corrected DUT results
    Protected Friend myDUTMeasurementPaths() As String       'The strings with the DUT locations. These are used to create the subdirectories for the corrected DUT results
    Protected Friend InternalCrosstalkModel As Boolean = False  'Set this true to use the internal crosstalk model rather than the conventional crosstalk model.

    'The before and after calibration models
    Private myBeforeCalibrationModels() As Object
    Private myBeforeCalibrationModelNames() As String
    Private myBeforeCalibrationStdTypes() As String
    Private myBeforeCalibrationLengths() As Object
    Private myBeforeCalibrationMeasurements() As Object

    Private myBeforePMMeasurementW1P As Object
    Private myBeforeHPRMeasurementW1P As Object
    Private myBeforePMAdapterS2P As Object
    Private myBeforeHPRAdapterS2P As Object
    Private myBeforePMMismatchS2P As Object
    Private myBeforeHPRMismatchS2P As Object
    Private myBeforePMW1P As Object
    Private myBeforeHPRW1P As Object

    Private myAfterCalibrationModels() As Object
    Private myAfterCalibrationModelNames() As String
    Private myAfterCalibrationStdTypes() As String
    Private myAfterCalibrationLengths() As Object
    Private myAfterCalibrationMeasurements() As Object
    Private myBeforePort1Connection() As Integer
    Private myBeforePort2Connection() As Integer
    Private myAfterPort1Connection() As Integer
    Private myAfterPort2Connection() As Integer

    Private myAfterPMMeasurementW1P As Object
    Private myAfterHPRMeasurementW1P As Object
    Private myAfterPMAdapterS2P As Object
    Private myAfterHPRAdapterS2P As Object
    Private myAfterPMMismatchS2P As Object
    Private myAfterHPRMismatchS2P As Object
    Private myAfterPMW1P As Object
    Private myAfterHPRW1P As Object

    Public Sub New()

    End Sub

    ''' <summary>
    ''' Calibrate the DUTs after the correction is performed.
    ''' The calibration error model in the mechanism list is used.
    ''' </summary>
    ''' <param name="SimulationCount">Automatically saves the version for this particular simulation</param>
    ''' <param name="myFinalResultsPath"></param>
    ''' <param name="MechanismList1">The mechanism list</param>
    ''' <param name="UnpreturbedSolution">Set true for the unpreturbed solution.</param>
    ''' <param name="DUTSuffix">The DUT suffix that get added to dut names.</param>
    ''' <param name="Binary">Determines whether results are also saved in binary format.</param>
    ''' <param name="FillDCwxpValues">Read in DC values from file and add to .wxp output file.</param>
    ''' <param name="DUTRawMeasurementPaths">Paths to DC values from file for .wxp output file.</param>
    ''' <remarks></remarks>
    Public Sub SaveCalibratedDUTs(ByVal SimulationCount As Integer, ByVal myFinalResultsPath As String, ByVal MechanismList1 As MechanismList, ByVal UnpreturbedSolution As Boolean, ByVal DUTSuffix As String, ByVal Binary As Boolean, Optional ByVal FillDCwxpValues As Boolean = False, Optional ByVal DUTRawMeasurementPaths() As String = Nothing)

        'Save the new terms to disk
        Dim BinaryExtention As String = "" : If Binary Then BinaryExtention = "_binary"
        Dim PreSolutionNPorts As Integer = MechanismList1.PreSolution.NPorts    'This would be 2*2=4 for a 2-port calibration
        Dim DUTNPorts As Integer = PreSolutionNPorts / 2
        If 2 * DUTNPorts <> PreSolutionNPorts Then Throw New ApplicationException("CalibrationEngine SaveCalibratedDUTs: Number of ports incorrect.")
        If Not My.Computer.FileSystem.DirectoryExists(myFinalResultsPath) Then My.Computer.FileSystem.CreateDirectory(myFinalResultsPath)

        'Write out switch-terms for the solution
        MechanismList1.SwitchTerms.Write(myFinalResultsPath + "\SwitchTerms_" + SimulationCount.ToString + ".switch")
        If Binary Then MechanismList1.SwitchTerms.Write(myFinalResultsPath + "\SwitchTerms_" + SimulationCount.ToString + ".switch" + BinaryExtention)
        If DUTNPorts = 2 Then MechanismList1.IsolationTerms.Write(myFinalResultsPath + "\IsolationTerms_" + SimulationCount.ToString + ".iso")
        If Binary Then If DUTNPorts = 2 Then MechanismList1.IsolationTerms.Write(myFinalResultsPath + "\IsolationTerms_" + SimulationCount.ToString + ".iso" + BinaryExtention)

        'Write out the calibration scattering-parameter matrix and efective dielectric constant.
        MechanismList1.PreSolution.Write(myFinalResultsPath + "\Solution_" + SimulationCount.ToString + ".s" + PreSolutionNPorts.ToString + "p")
        If Binary Then MechanismList1.PreSolution.Write(myFinalResultsPath + "\Solution_" + SimulationCount.ToString + ".s" + PreSolutionNPorts.ToString + "p" + BinaryExtention)
        MechanismList1.EpsEff.Write(myFinalResultsPath + "\EPSr_" + SimulationCount.ToString + ".complex")
        If Binary Then MechanismList1.EpsEff.Write(myFinalResultsPath + "\EPSr_" + SimulationCount.ToString + ".complex" + BinaryExtention)

        'Figure out where the corrected scattering-parameters will go.
        Dim mySubPath As String = Left(myFinalResultsPath, InStrRev(myFinalResultsPath, "\") - 1)    'This is the subpath (Presolution, Covariance, or MonteCarlo) for the DUT data
        mySubPath = Mid(mySubPath, InStrRev(mySubPath, "\"))    'This is the subpath (Presolution, Covariance, or MonteCarlo) for the DUT data
        'If (InStr(mySubPath, "Covariance") + InStr(mySubPath, "MonteCarlo") + InStr(mySubPath, "PreSolution") = 0) And SimulationCount = 0 Then mySubPath = ""
        If UnpreturbedSolution Then mySubPath = ""

        'Correct the DUTs and store the results
        Dim NDUT As Integer = myDUTMeasurementNames.Count
        If NDUT > 0 Then
            For i As Integer = 0 To NDUT - 1

                'Create a folder for the results if it does not already exist.
                Dim myFinalResultsPathDUT As String = System.IO.Path.GetDirectoryName(myDUTMeasurementPaths(i)) + "\" + myDUTMeasurementNames(i) + DUTSuffix + "_Support" + mySubPath
                If Not My.Computer.FileSystem.DirectoryExists(myFinalResultsPathDUT) Then My.Computer.FileSystem.CreateDirectory(myFinalResultsPathDUT)

                'Figure out if these are wave or scattering parameters.
                Dim IsWnP As Boolean = Left(System.IO.Path.GetExtension(myDUTMeasurementPaths(i)), 2).ToLower = ".w"
                Dim FileNameSaveDUT As String
                If IsWnP Then   '.wnp file

                    'Correct the results with the solution FileNameSolution and save them out to disk as FileNameSaveDUT
                    FileNameSaveDUT = myFinalResultsPathDUT + "\" + myDUTMeasurementNames(i) + DUTSuffix + "_" + SimulationCount.ToString + ".w" + DUTNPorts.ToString + "p"
                    Dim DUT_WParams As New RealMatrix(MechanismList1.FrequencyList.NRows, 1 + 4 * DUTNPorts * DUTNPorts, ".w" + DUTNPorts.ToString + "p")

                    'Get the scattering parameters of the DUT.
                    Dim Temp As New RealMatrix(MechanismList1.FrequencyList.NRows, 1 + 4 * DUTNPorts * DUTNPorts, ".w" + DUTNPorts.ToString + "p")
                    Temp = myDUTMeasurements(i).getSParams(MechanismList1)
                    DUT_WParams.Fill(Temp)

                    'Calibrate should recognize the .wnp file type and react accordingly.
                    'The program should be smart enough to coorect propoerly for these if they are scattering parameters or wave files.
                    DUT_WParams.Calibrate(MechanismList1, InternalCrosstalkModel)

                    'Write out voltages and currents as well if they are there.
                    Dim AuxDCFilePath As String = ""
                    If Not IsNothing(DUTRawMeasurementPaths) Then AuxDCFilePath = System.IO.Path.GetDirectoryName(DUTRawMeasurementPaths(i)) + "\" + System.IO.Path.GetFileNameWithoutExtension(DUTRawMeasurementPaths(i)) + ".DCw" + DUTNPorts.ToString + "p"
                    If FillDCwxpValues And My.Computer.FileSystem.FileExists(AuxDCFilePath) Then  'DC voltages and currents are there.

                        'Add in first line with DC values
                        Dim DUT_DCvi As New RealMatrix(1, 1 + 4 * DUTNPorts * DUTNPorts, ".w" + DUTNPorts.ToString + "p")
                        DUT_DCvi.Read(AuxDCFilePath)
                        Dim DUT_WParamsWithDCvi As New RealMatrix(MechanismList1.FrequencyList.NRows + DUT_DCvi.NRows, 1 + 4 * DUTNPorts * DUTNPorts, ".w" + DUTNPorts.ToString + "p")
                        For k1 As Integer = 1 To DUT_DCvi.NRows
                            For k2 As Integer = 1 To 1 + 4 * DUTNPorts * DUTNPorts
                                DUT_WParamsWithDCvi(k1, k2) = DUT_DCvi(k1, k2)
                            Next k2
                        Next k1
                        'Add in the rest of the Wave Parameters
                        For k1 As Integer = 1 To MechanismList1.FrequencyList.NRows
                            For k2 As Integer = 1 To 1 + 4 * DUTNPorts * DUTNPorts
                                DUT_WParamsWithDCvi(k1 + DUT_DCvi.NRows, k2) = DUT_WParams(k1, k2)
                            Next k2
                        Next k1
                        'Write out the result with the DC terms added.
                        DUT_WParamsWithDCvi.Write(FileNameSaveDUT + BinaryExtention)

                    Else    'Write out the normal result

                        DUT_WParams.Write(FileNameSaveDUT + BinaryExtention)

                    End If

                    'Write out the .snp file as well.
                    Dim DUT_SParams As RealMatrix
                    DUT_SParams = DUT_WParams.WnP_to_SnP    'CustomFormControls.PostProcessorModule.WnP_to_SnP(DUT_WParams)
                    Dim FileNameSaveDUT_SPar As String = System.IO.Path.GetDirectoryName(FileNameSaveDUT) + "\" + System.IO.Path.GetFileNameWithoutExtension(FileNameSaveDUT) + ".s" + DUTNPorts.ToString + "p"
                    'DUT_SParams.Write(FileNameSaveDUT_SPar + BinaryExtention)

                    'Write out voltages and currents as well if they are there.

                    AuxDCFilePath = ""
                    If Not IsNothing(DUTRawMeasurementPaths) Then AuxDCFilePath = System.IO.Path.GetDirectoryName(DUTRawMeasurementPaths(i)) + "\" + System.IO.Path.GetFileNameWithoutExtension(DUTRawMeasurementPaths(i)) + ".DCs" + DUTNPorts.ToString + "p"
                    If FillDCwxpValues And My.Computer.FileSystem.FileExists(AuxDCFilePath) Then  'DC voltages and currents are there.

                        'Add in first line with DC values
                        Dim DUT_DCvi As New RealMatrix(1, 1 + 2 * DUTNPorts * DUTNPorts, ".w" + DUTNPorts.ToString + "p")
                        DUT_DCvi.Read(AuxDCFilePath)
                        Dim DUT_SParamsWithDCvi As New RealMatrix(MechanismList1.FrequencyList.NRows + +DUT_DCvi.NRows, 1 + 2 * DUTNPorts * DUTNPorts, ".s" + DUTNPorts.ToString + "p")
                        For k1 As Integer = 1 To DUT_DCvi.NRows
                            For k2 As Integer = 1 To 1 + 2 * DUTNPorts * DUTNPorts
                                DUT_SParamsWithDCvi(k1, k2) = DUT_DCvi(k1, k2)
                            Next k2
                        Next k1
                        'Add in the rest of the Wave Parameters
                        For k1 As Integer = 1 To MechanismList1.FrequencyList.NRows
                            For k2 As Integer = 1 To 1 + 2 * DUTNPorts * DUTNPorts
                                DUT_SParamsWithDCvi(k1 + DUT_DCvi.NRows, k2) = DUT_SParams(k1, k2)
                            Next k2
                        Next k1
                        'Write out the result with the DC terms added.
                        DUT_SParamsWithDCvi.Write(FileNameSaveDUT_SPar + BinaryExtention)

                    Else    'Write out the normal result

                        DUT_SParams.Write(FileNameSaveDUT_SPar + BinaryExtention)

                    End If

                Else            '.snp file

                    'Correct the results with the solution FileNameSolution and save them out to disk as FileNameSaveDUT
                    FileNameSaveDUT = myFinalResultsPathDUT + "\" + myDUTMeasurementNames(i) + DUTSuffix + "_" + SimulationCount.ToString + ".s" + DUTNPorts.ToString + "p"
                    Dim DUT_SParams As New RealMatrix(MechanismList1.FrequencyList.NRows, 1 + 2 * DUTNPorts * DUTNPorts, ".s" + DUTNPorts.ToString + "p")

                    'Get the scattering parameters of the DUT.
                    Dim Temp As New RealMatrix(MechanismList1.FrequencyList.NRows, 1 + 2 * DUTNPorts * DUTNPorts, ".s" + DUTNPorts.ToString + "p")
                    Temp = myDUTMeasurements(i).getSParams(MechanismList1)
                    DUT_SParams.Fill(Temp)

                    'Calibrate should recognize the .wnp file type and react accordingly.
                    'The program should be smart enough to coorect propoerly for these if they are scattering parameters or wave files.
                    DUT_SParams.Calibrate(MechanismList1, InternalCrosstalkModel)
                    'DUT_SParams.Write(FileNameSaveDUT + BinaryExtention)

                    'Write out voltages and currents as well if they are there.

                    Dim AuxDCFilePath As String = ""
                    If Not IsNothing(DUTRawMeasurementPaths) Then AuxDCFilePath = System.IO.Path.GetDirectoryName(DUTRawMeasurementPaths(i)) + "\" + System.IO.Path.GetFileNameWithoutExtension(DUTRawMeasurementPaths(i)) + ".DCs" + DUTNPorts.ToString + "p"
                    If FillDCwxpValues And My.Computer.FileSystem.FileExists(AuxDCFilePath) Then  'DC voltages and currents are there.

                        'Add in first line with DC values
                        Dim DUT_DCvi As New RealMatrix(1, 1 + 2 * DUTNPorts * DUTNPorts, ".w" + DUTNPorts.ToString + "p")
                        DUT_DCvi.Read(AuxDCFilePath)
                        Dim DUT_SParamsWithDCvi As New RealMatrix(MechanismList1.FrequencyList.NRows + +DUT_DCvi.NRows, 1 + 2 * DUTNPorts * DUTNPorts, ".s" + DUTNPorts.ToString + "p")
                        For k1 As Integer = 1 To DUT_DCvi.NRows
                            For k2 As Integer = 1 To 1 + 2 * DUTNPorts * DUTNPorts
                                DUT_SParamsWithDCvi(k1, k2) = DUT_DCvi(k1, k2)
                            Next k2
                        Next k1
                        'Add in the rest of the Wave Parameters
                        For k1 As Integer = 1 To MechanismList1.FrequencyList.NRows
                            For k2 As Integer = 1 To 1 + 2 * DUTNPorts * DUTNPorts
                                DUT_SParamsWithDCvi(k1 + DUT_DCvi.NRows, k2) = DUT_SParams(k1, k2)
                            Next k2
                        Next k1
                        'Write out the result with the DC terms added.
                        DUT_SParamsWithDCvi.Write(FileNameSaveDUT + BinaryExtention)

                    Else    'Write out the normal result

                        DUT_SParams.Write(FileNameSaveDUT + BinaryExtention)

                    End If

                End If


            Next
        End If

    End Sub

    Public Sub SmoothErrorBoxes(ByRef MechanismList1 As MechanismList, ByVal TurnOnSmoothingValue As Double)

        'Set up Tpre, the original unperturbed solution, Tnew, the perturbed solution, and all of the quadrants
        Dim Tpre As New ComplexMatrix(4, 4), Tnew As New ComplexMatrix(4, 4)
        Dim TpreEC As New ComplexMatrix(4, 4), TnewEC As New ComplexMatrix(4, 4)
        Dim TpreCP As New ComplexMatrix(4, 4), TnewCP As New ComplexMatrix(4, 4)
        Dim Tpre_11 As New ComplexMatrix(2, 2), Tnew_11 As New ComplexMatrix(2, 2)
        Dim Tpre_21 As New ComplexMatrix(2, 2), Tnew_21 As New ComplexMatrix(2, 2)
        Dim Tpre_12 As New ComplexMatrix(2, 2), Tnew_12 As New ComplexMatrix(2, 2)
        Dim Tpre_22 As New ComplexMatrix(2, 2), Tnew_22 As New ComplexMatrix(2, 2)

        'Set up Spre, the original unperturbed solution, Snew, the perturbed solution, and all of the quadrants
        Dim Spre As New ComplexMatrix(4, 4), Snew As New ComplexMatrix(4, 4)
        Dim SpreEC As New ComplexMatrix(4, 4), SnewEC As New ComplexMatrix(4, 4)
        Dim SpreCP As New ComplexMatrix(4, 4), SnewCP As New ComplexMatrix(4, 4)
        Dim Spre_11 As New ComplexMatrix(2, 2), Snew_11 As New ComplexMatrix(2, 2)
        Dim Spre_21 As New ComplexMatrix(2, 2), Snew_21 As New ComplexMatrix(2, 2)
        Dim Spre_12 As New ComplexMatrix(2, 2), Snew_12 As New ComplexMatrix(2, 2)
        Dim Spre_22 As New ComplexMatrix(2, 2), Snew_22 As New ComplexMatrix(2, 2)

        'The perturbations, flags, etc.
        Dim DeltaT(MechanismList1.FrequencyList.NRows) As ComplexMatrix
        Dim DeltaTJump(MechanismList1.FrequencyList.NRows) As Boolean

        'Calculate the pertubations
        For i As Integer = 1 To MechanismList1.FrequencyList.NRows                      'Step through the frequencies

            'Get the old and new S-parameter solution
            Spre.Fill(MechanismList1.PreSolution_Cashed.SMatrix(i))
            Snew.Fill(MechanismList1.PreSolution.SMatrix(i))

            'Transform to T parameters
            If InternalCrosstalkModel Then  'The more complex two-tier case.
                'The Error Coefficients of the VNA
                Call RealMatrix.GetErrorBoxes(Spre, Spre_11, Spre_12, Spre_21, Spre_22)
                SpreEC = ReAssemble4x4(Spre_11, Spre_12, Spre_21, Spre_22)
                TpreEC = TFromS4(SpreEC)
                'The coupling coefficients of the VNA
                Call RealMatrix.GetCouplingTerms(Spre, Spre_11, Spre_12, Spre_21, Spre_22)
                SpreCP = ReAssemble4x4(Spre_11, Spre_12, Spre_21, Spre_22)
                TpreCP = TFromS4(SpreCP)
                'The Error Coefficients of the VNA
                Call RealMatrix.GetErrorBoxes(Snew, Snew_11, Snew_12, Snew_21, Snew_22)
                SnewEC = ReAssemble4x4(Snew_11, Snew_12, Snew_21, Snew_22)
                TnewEC = TFromS4(SnewEC)
                'The coupling coefficients of the VNA
                Call RealMatrix.GetCouplingTerms(Snew, Snew_11, Snew_12, Snew_21, Snew_22)
                SnewCP = ReAssemble4x4(Snew_11, Snew_12, Snew_21, Snew_22)
                TnewCP = TFromS4(SnewCP)
                'Find DeltaT
                DeltaT(i) = (TpreCP ^ -1) * (TpreEC ^ -1) * TnewEC * TnewCP
            Else    'The conventional case.
                Tpre = TFromS4(Spre) : Tnew = TFromS4(Snew)
                DeltaT(i) = (Tpre ^ -1) * Tnew
            End If

        Next i

        'Look for jumps
        For i As Integer = 1 To MechanismList1.FrequencyList.NRows                      'Step through the frequencies
            Dim Average As New ComplexMatrix(4, 4), Istart As Integer = i - 2, Icount As Integer = 0
            If Istart < 1 Then Istart = 1
            For k As Integer = Istart To i
                'Round up all of the good nearby candidates for good results to average together
                If k <> i And Not DeltaTJump(k) Then
                    Average = Average + DeltaT(k) : Icount = Icount + 1
                End If
            Next k
            'See if we have a big jump or not
            If Icount > 0 Then
                Dim Difference As New ComplexMatrix(4, 4), SumDeltas As Double = 0.0
                Difference = DeltaT(i) - (1.0 / CDbl(Icount)) * Average
                For k1 As Integer = 1 To 4
                    For k2 As Integer = 1 To 4
                        SumDeltas = SumDeltas + Abs(Difference(k1, k2))
                    Next k2
                Next k1
                DeltaTJump(i) = (SumDeltas > TurnOnSmoothingValue)   'Flag any big jumps
            End If
        Next i

        'Interpolate any big jumps
        For i As Integer = 1 To MechanismList1.FrequencyList.NRows                      'Step through the frequencies
            If DeltaTJump(i) Then   'Found a big jump.
                Dim Average As New ComplexMatrix(4, 4), ER_Average As Double = 0.0, EI_Average As Double = 0.0
                Dim Istart As Integer = i - 2, Istop As Integer = i + 2, Icount As Integer = 0
                If Istart < 1 Then Istart = 1
                If Istop > MechanismList1.FrequencyList.NRows Then Istop = MechanismList1.FrequencyList.NRows
                For k As Integer = Istart To Istop
                    If k <> i And Not DeltaTJump(k) Then
                        Average = Average + DeltaT(k)
                        ER_Average = ER_Average + MechanismList1.EpsEff(k, 2)
                        EI_Average = EI_Average + MechanismList1.EpsEff(k, 3)
                        Icount = Icount + 1
                    End If
                Next k
                'Try to fix the jump
                If Icount > 0 Then
                    Dim Difference As New ComplexMatrix(4, 4), SumDeltas As Double = 0.0
                    DeltaT(i) = (1.0 / CDbl(Icount)) * Average  'The new DeltaT based on the nearby good values

                    'Get the old pre-solution back
                    Spre.Fill(MechanismList1.PreSolution_Cashed.SMatrix(i))

                    'Use Tpre and the nearby values for DeltaT to update Tnew ~ Tpre * DeltaT
                    If InternalCrosstalkModel Then
                        'The Error Coefficients of the VNA
                        Call RealMatrix.GetErrorBoxes(Spre, Spre_11, Spre_12, Spre_21, Spre_22)
                        SpreEC = ReAssemble4x4(Spre_11, Spre_12, Spre_21, Spre_22)
                        TpreEC = TFromS4(SpreEC)
                        'The coupling coefficients of the VNA
                        Call RealMatrix.GetCouplingTerms(Spre, Spre_11, Spre_12, Spre_21, Spre_22)
                        SpreCP = ReAssemble4x4(Spre_11, Spre_12, Spre_21, Spre_22)
                        TpreCP = TFromS4(SpreCP)
                        Tnew = TpreEC * TpreCP * DeltaT(i)
                    Else
                        Spre.Fill(MechanismList1.PreSolution_Cashed.SMatrix(i))
                        Tpre = TFromS4(Spre)
                        Tnew = Tpre * DeltaT(i)
                    End If

                    'Save the data back into the Mechanism List
                    Snew = SFromT4(Tnew)
                    MechanismList1.PreSolution_SMatrix(i) = Snew
                    MechanismList1.EpsEff(i, 2) = (1.0 / CDbl(Icount)) * ER_Average
                    MechanismList1.EpsEff(i, 3) = (1.0 / CDbl(Icount)) * EI_Average

                End If
            End If
        Next i

    End Sub
    ''' <summary>
    ''' Reassemble a 2nx2n matrix from its nxn quadrants
    ''' </summary>
    ''' <param name="S11"></param>
    ''' <param name="S12"></param>
    ''' <param name="S21"></param>
    ''' <param name="S22"></param>
    ''' <returns>The reassembled 4x4 complex matrix</returns>
    ''' <remarks></remarks>
    Private Function ReAssemble4x4(ByVal S11 As ComplexMatrix, ByVal S12 As ComplexMatrix, ByVal S21 As ComplexMatrix, ByVal S22 As ComplexMatrix) As ComplexMatrix

        Dim N As Integer = S11.NRows : If N <> S11.NCols Then Throw New ApplicationException("CalibrationEngine ReAssemble4x4: Error in dimensions")
        Dim S As New ComplexMatrix(2 * N, 2 * N)

        'Reassemble the 2nx2n scattering-parameter matrix
        For k1 As Integer = 1 To N
            For k2 As Integer = 1 To N
                S(k1, k2) = S11(k1, k2) : S(k1, k2 + N) = S12(k1, k2)
                S(k1 + N, k2) = S21(k1, k2) : S(k1 + N, k2 + N) = S22(k1, k2)
            Next k2
        Next k1

        Return S

    End Function
    ''' <summary>
    ''' Return the scattering parameters of a 4x4 complex transmission matrix
    ''' </summary>
    ''' <param name="T">The input 4x4 transmission matrix</param>
    ''' <returns>The 4x4 scattering-parameter matrix.</returns>
    ''' <remarks></remarks>
    Private Function SFromT4(ByVal T As ComplexMatrix) As ComplexMatrix

        Dim T11 As New ComplexMatrix(2, 2), T21 As New ComplexMatrix(2, 2), T12 As New ComplexMatrix(2, 2), T22 As New ComplexMatrix(2, 2)

        'Split the scattering parameters into quadrants
        Call RealMatrix.GetSplitParts(T, T11, T12, T21, T22)

        'Set up the matrices and dimensions.
        Dim N As Integer = T11.NRows
        Dim S11 As New ComplexMatrix(N, N), S21 As New ComplexMatrix(N, N), S12 As New ComplexMatrix(N, N), S22 As New ComplexMatrix(N, N)


        'Calculate the transmission parameters of the quadrants from the scattering parameters of the quadrants
        S21 = T22 ^ -1 : S22 = -S21 * T21 : S11 = T12 * S21 : S12 = T11 - S11 * T21

        Return ReAssemble4x4(S11, S12, S21, S22)

    End Function
    ''' <summary>
    ''' Return the tranmission parameters of a 4x4 complex scattering-parameter matrix
    ''' </summary>
    ''' <param name="S">The input 4x4 scattering-parameter matrix</param>
    ''' <returns>The 4x4 transmission matrix.</returns>
    ''' <remarks></remarks>
    Private Function TFromS4(ByVal S As ComplexMatrix) As ComplexMatrix

        'Set up the 2x2 quadrants for the 4x4 matrices
        Dim S11 As New ComplexMatrix(2, 2), S21 As New ComplexMatrix(2, 2), S12 As New ComplexMatrix(2, 2), S22 As New ComplexMatrix(2, 2)

        'Split the scattering parameters into quadrants
        Call RealMatrix.GetSplitParts(S, S11, S12, S21, S22)

        'Set up the matrices and dimensions.
        Dim N As Integer = S11.NRows
        Dim T11 As New ComplexMatrix(N, N), T21 As New ComplexMatrix(N, N), T12 As New ComplexMatrix(N, N), T22 As New ComplexMatrix(N, N)

        'Calculate the transmission parameters of the quadrants from the scattering parameters of the quadrants
        T22 = S21 ^ -1 : T21 = -T22 * S22 : T12 = S11 * T22 : T11 = S12 - T12 * S22

        Return ReAssemble4x4(T11, T12, T21, T22)

    End Function

    ''' <summary>
    ''' Gets set by the caller. Can be accessed by the calibration algorithm if needed.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property NumberOfPorts() As Integer
        Get
            Return myNumberOfPorts
        End Get
        Set(ByVal value As Integer)
            myNumberOfPorts = value
        End Set
    End Property

    ''' <summary>
    ''' The ports that the calibration standards are connected to
    ''' </summary>
    ''' <value></value>
    ''' <returns>The port that port 1 is connected to.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Port1Connection() As Integer()
        Get
            Return myPort1Connection
        End Get
    End Property
    ''' <summary>
    ''' The ports that the calibration standards are connected to
    ''' </summary>
    ''' <value></value>
    ''' <returns>The port that port 2 is connected to.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Port2Connection() As Integer()
        Get
            Return myPort2Connection
        End Get
    End Property

    ''' <summary>
    ''' Initializes the calibration engine and sets up the objects needed to do the calibration
    ''' </summary>
    ''' <param name="PortModels">The port parameters Cable1, TestPort1, DUTPort1, Cable2, TestPort2, DUTPort2</param>
    ''' <param name="BeforeCalibrationModels">The Before calibration definitions</param>
    ''' <param name="BeforeCalibrationModelNames"></param>
    ''' <param name="BeforeCalibrationStdTypes"></param>
    ''' <param name="BeforeCalibrationLengths"></param>
    ''' <param name="BeforeCalibrationMeasurements"></param>
    ''' <param name="AfterCalibrationModels"></param>
    ''' <param name="AfterCalibrationModelNames"></param>
    ''' <param name="AfterCalibrationStdTypes"></param>
    ''' <param name="AfterCalibrationLengths"></param>
    ''' <param name="AfterCalibrationMeasurements"></param>
    ''' <param name="DUTMeasurements"></param>
    ''' <param name="DUTMeasurementNames"></param>
    ''' <param name="DUTMeasurementPaths"></param>
    ''' <param name="BeforePort1Connection"></param>
    ''' <param name="BeforePort2Connection"></param>
    ''' <param name="AfterPort1Connection"></param>
    ''' <param name="AfterPort2Connection"></param>
    ''' <remarks>Each of these can be called with .getSParams(MechanismList) to return scattering parameters.</remarks>
    Public Sub InitializeModels(ByVal PortModels() As Object, ByVal DUTPortModels() As Object, ByVal BeforeCalibrationModels() As Object, ByVal BeforeCalibrationModelNames() As String, ByVal BeforeCalibrationStdTypes() As String, ByVal BeforeCalibrationLengths() As Object, ByVal BeforeCalibrationMeasurements() As Object, ByVal AfterCalibrationModels() As Object, ByVal AfterCalibrationModelNames() As String, ByVal AfterCalibrationStdTypes() As String, ByVal AfterCalibrationLengths() As Object, ByVal AfterCalibrationMeasurements() As Object, ByVal DUTMeasurements() As Object, ByVal DUTMeasurementNames() As String, ByVal DUTMeasurementPaths() As String, ByVal BeforePort1Connection() As Integer, ByVal BeforePort2Connection() As Integer, ByVal AfterPort1Connection() As Integer, ByVal AfterPort2Connection() As Integer)

        myPortModels = PortModels
        myDUTPortModels = DUTPortModels
        myBeforeCalibrationModels = BeforeCalibrationModels
        myBeforeCalibrationModelNames = BeforeCalibrationModelNames
        myBeforeCalibrationStdTypes = BeforeCalibrationStdTypes
        myBeforeCalibrationLengths = BeforeCalibrationLengths
        myBeforeCalibrationMeasurements = BeforeCalibrationMeasurements
        myAfterCalibrationModels = AfterCalibrationModels
        myAfterCalibrationModelNames = AfterCalibrationModelNames
        myAfterCalibrationStdTypes = AfterCalibrationStdTypes
        myAfterCalibrationLengths = AfterCalibrationLengths
        myAfterCalibrationMeasurements = AfterCalibrationMeasurements
        myDUTMeasurements = DUTMeasurements
        myDUTMeasurementNames = DUTMeasurementNames
        myDUTMeasurementPaths = DUTMeasurementPaths
        myBeforePort1Connection = BeforePort1Connection
        myBeforePort2Connection = BeforePort2Connection
        myAfterPort1Connection = AfterPort1Connection
        myAfterPort2Connection = AfterPort2Connection

    End Sub

    ''' <summary>
    ''' Initializes the LSNA power-phase calibration components.
    ''' </summary>
    ''' <param name="BeforePMMeasurementW1P"></param>
    ''' <param name="BeforeHPRMeasurementW1P"></param>
    ''' <param name="BeforePMAdapterS2P"></param>
    ''' <param name="BeforeHPRAdapterS2P"></param>
    ''' <param name="BeforePMMismatchS2P"></param>
    ''' <param name="BeforeHPRMismatchS2P"></param>
    ''' <param name="BeforePMW1P"></param>
    ''' <param name="BeforeHPRW1P"></param>
    ''' <param name="AfterPMMeasurementW1P"></param>
    ''' <param name="AfterHPRMeasurementW1P"></param>
    ''' <param name="AfterPMAdapterS2P"></param>
    ''' <param name="AfterHPRAdapterS2P"></param>
    ''' <param name="AfterPMMismatchS2P"></param>
    ''' <param name="AfterHPRMismatchS2P"></param>
    ''' <param name="AfterPMW1P"></param>
    ''' <param name="AfterHPRW1P"></param>
    ''' <remarks></remarks>
    Public Sub InitializeLSNAModels(ByVal BeforePMMeasurementW1P As Object, ByVal BeforeHPRMeasurementW1P As Object, ByVal BeforePMAdapterS2P As Object, ByVal BeforeHPRAdapterS2P As Object, ByVal BeforePMMismatchS2P As Object, ByVal BeforeHPRMismatchS2P As Object, ByVal BeforePMW1P As Object, ByVal BeforeHPRW1P As Object, ByVal AfterPMMeasurementW1P As Object, ByVal AfterHPRMeasurementW1P As Object, ByVal AfterPMAdapterS2P As Object, ByVal AfterHPRAdapterS2P As Object, ByVal AfterPMMismatchS2P As Object, ByVal AfterHPRMismatchS2P As Object, ByVal AfterPMW1P As Object, ByVal AfterHPRW1P As Object)

        myBeforePMMeasurementW1P = BeforePMMeasurementW1P
        myBeforePMAdapterS2P = BeforePMAdapterS2P
        myBeforePMMismatchS2P = BeforePMMismatchS2P
        myBeforePMW1P = BeforePMW1P
        myBeforeHPRMeasurementW1P = BeforeHPRMeasurementW1P
        myBeforeHPRAdapterS2P = BeforeHPRAdapterS2P
        myBeforeHPRMismatchS2P = BeforeHPRMismatchS2P
        myBeforeHPRW1P = BeforeHPRW1P

        myAfterPMMeasurementW1P = AfterPMMeasurementW1P
        myAfterPMAdapterS2P = AfterPMAdapterS2P
        myAfterPMMismatchS2P = AfterPMMismatchS2P
        myAfterPMW1P = AfterPMW1P
        myAfterHPRMeasurementW1P = AfterHPRMeasurementW1P
        myAfterHPRAdapterS2P = AfterHPRAdapterS2P
        myAfterHPRMismatchS2P = AfterHPRMismatchS2P
        myAfterHPRW1P = AfterHPRW1P

    End Sub


    ''' <summary>
    ''' Get the 2x2 port 1 calibration matrix from the Mechanism List.
    ''' </summary>
    ''' <param name="MechanismList1"></param>
    ''' <param name="Port_SelectedIndex">Determines which port to get calibration matrix from</param>
    ''' <returns>.s2p format one-port calibration matrix.</returns>
    ''' <remarks>Should work no matter what the number of ports is.</remarks>
    Private Function GetPortCalibration(ByVal MechanismList1 As MechanismList, ByVal Port_SelectedIndex As Integer) As RealMatrix

        'Pull out the calibration matrix from MechanismList1.PreSolution and get the port-1 s-parameter matrix.
        Dim SParCalibration As RealMatrix = MechanismList1.PreSolution
        Dim SPArCalibrationSubMatrix As New ComplexMatrix(SParCalibration.NPorts, SParCalibration.NPorts)
        Dim NPortsVNA As Integer = SParCalibration.NPorts / 2
        If 2 * NPortsVNA <> SParCalibration.NPorts Then Throw New ApplicationException("CalibrationEngine:GetPort1Calibration: Number of ports is incorrect.")
        Dim SParCalibrationPort1 As New RealMatrix(MechanismList1.FrequencyList.NRows, 9, ".s2p")
        SParCalibrationPort1.Vector(1) = MechanismList1.FrequencyList
        Dim SParCalibrationPort1SubMatrix As New ComplexMatrix(2, 2)
        For i As Integer = 1 To MechanismList1.FrequencyList.NRows
            SPArCalibrationSubMatrix = SParCalibration.SMatrix(i)
            SParCalibrationPort1SubMatrix(1, 1) = SPArCalibrationSubMatrix(1 + Port_SelectedIndex, 1 + Port_SelectedIndex)
            SParCalibrationPort1SubMatrix(1, 2) = SPArCalibrationSubMatrix(1 + Port_SelectedIndex, 1 + Port_SelectedIndex + NPortsVNA)
            SParCalibrationPort1SubMatrix(2, 1) = SPArCalibrationSubMatrix(1 + Port_SelectedIndex + NPortsVNA, 1 + Port_SelectedIndex)
            SParCalibrationPort1SubMatrix(2, 2) = SPArCalibrationSubMatrix(1 + Port_SelectedIndex + NPortsVNA, 1 + Port_SelectedIndex + NPortsVNA)
            SParCalibrationPort1.SMatrix(i) = SParCalibrationPort1SubMatrix
        Next i

        Return SParCalibrationPort1

    End Function

    ''' <summary>
    ''' Calculate the ratio of the power the LSNA thinks that it generated to the power that was actually generated
    ''' </summary>
    ''' <param name="MechanismList1"></param>
    ''' <param name="PM_Port_SelectedIndex">Port to which the power meter is connected</param>
    ''' <returns>The ratio of the power the LSNA thinks that it generated to the power that was actually generated</returns>
    ''' <remarks></remarks>
    Public Function PowerCalRatio(ByVal MechanismList1 As MechanismList, ByVal PM_Port_SelectedIndex As Integer) As RealMatrix

        'Find the power calibration ratio for the before calibration
        Dim PowerCalRatio1 As New RealMatrix(MechanismList1.FrequencyList.NRows, 3, ".complex")
        PowerCalRatio1.Vector(1) = MechanismList1.FrequencyList

        'Pull out the calibration matrix from MechanismList1.PreSolution and get the port-1 s-parameter matrix.
        Dim SParCalibrationPort1 As New RealMatrix(MechanismList1.FrequencyList.NRows, 9, ".s2p")
        SParCalibrationPort1.Fill(GetPortCalibration(MechanismList1, PM_Port_SelectedIndex))

        'Find the scattering parameters of the adapter and mismatches
        Dim SPar As New RealMatrix(MechanismList1.FrequencyList.NRows, 9, ".s2p")
        SPar.InitializeAsSParams()
        SPar.Vector(1) = MechanismList1.FrequencyList
        Dim SPar1 As New RealMatrix(MechanismList1.FrequencyList.NRows, 9, ".s2p")
        SPar1.InitializeAsSParams()
        SPar1.Vector(1) = MechanismList1.FrequencyList
        Dim SPar2 As New RealMatrix(MechanismList1.FrequencyList.NRows, 9, ".s2p")
        SPar2.InitializeAsSParams()
        SPar2.Vector(1) = MechanismList1.FrequencyList
        SPar1 = myBeforePMAdapterS2P.getSParams(MechanismList1)
        SPar = myBeforePMMismatchS2P.getSParams(MechanismList1)
        If SPar(1, 4) = 0 And SPar(1, 5) = 0 And SPar(1, 6) = 0 And SPar(1, 7) = 0 Then 'User probably did not realize that transmission should be set to 1 
            For i As Integer = 1 To SPar.NRows
                SPar(i, 4) = 1.0 : SPar(i, 6) = 1.0
            Next i
        End If
        SPar2 = SPar1.CascadeSParameters(SPar)                  'Adapter*Mismatch
        SPar = SParCalibrationPort1.CascadeSParameters(SPar2)   'Calibration error box*Adapter*Mismatch.

        'The power b1 coming out of the LSNA and the power read by the power meter.
        Dim P_LSNA1 As New RealMatrix(MechanismList1.FrequencyList.NRows, 5, ".w1p")
        Dim P_Meter1 As New RealMatrix(MechanismList1.FrequencyList.NRows, 5, ".w1p")
        P_LSNA1 = myBeforePMMeasurementW1P.getSParams(MechanismList1)
        P_Meter1 = myBeforePMW1P.getSParams(MechanismList1)

        'Calculate the power calibration ratio for the before calibration
        For i As Integer = 1 To MechanismList1.FrequencyList.NRows

            'The raw measurement of the forward power a1 generated by the LSNA.
            Dim P_LSNA As New Complex(P_LSNA1(i, 2), P_LSNA1(i, 3))

            'Calculate the power b1 going into the power meter.
            P_LSNA = P_LSNA * toComplex(SPar(i, 4), SPar(i, 5))

            'Compare the power going into the power meter with the a1 power meter reads and form the ratio.
            P_LSNA = P_LSNA / toComplex(P_Meter1(i, 2), P_Meter1(i, 3))

            'Stuff into the output matrix
            PowerCalRatio1(i, 2) = P_LSNA.Re : PowerCalRatio1(i, 3) = P_LSNA.Im

        Next i

        'Find the power calibration ratio for the after calibration and average in the answers.
        If MechanismList1.BeforeAfterSwitch >= 0 And Not (myAfterPMMeasurementW1P Is Nothing) And Not (myAfterPMW1P Is Nothing) Then

            'Find the scattering parameters of the adapter and mismatches
            SPar.InitializeAsSParams()
            SPar1.InitializeAsSParams()
            SPar2.InitializeAsSParams()
            SPar1 = myAfterPMAdapterS2P.getSParams(MechanismList1)
            SPar = myAfterPMMismatchS2P.getSParams(MechanismList1)
            If SPar(1, 4) = 0 And SPar(1, 5) = 0 And SPar(1, 6) = 0 And SPar(1, 7) = 0 Then 'User probably did not realize that transmission should be set to 1 
                For i As Integer = 1 To SPar.NRows
                    SPar(i, 4) = 1.0 : SPar(i, 6) = 1.0
                Next i
            End If
            SPar2 = SPar1.CascadeSParameters(SPar)                  'Adapter*Mismatch
            SPar = SParCalibrationPort1.CascadeSParameters(SPar2)   'Calibration error box*Adapter*Mismatch.

            'The power b1 coming out of the LSNA and the power read by the power meter.
            P_LSNA1 = myAfterPMMeasurementW1P.getSParams(MechanismList1)
            P_Meter1 = myAfterPMW1P.getSParams(MechanismList1)

            'Calculate the power calibration ratio for the before calibration
            For i As Integer = 1 To MechanismList1.FrequencyList.NRows

                'Calculate the power b1 coming out of the LSNA.
                Dim P_LSNA As New Complex(P_LSNA1(i, 2), P_LSNA1(i, 3))

                'Calculate the power b1 going into the power meter.
                P_LSNA = P_LSNA * toComplex(SPar(i, 4), SPar(i, 5))

                'Compare the power going into the power meter with the actual a1 the power meter reads and form the ratio.
                P_LSNA = P_LSNA / toComplex(P_Meter1(i, 2), P_Meter1(i, 3))

                'Stuff into the output matrix, averaging the two results.
                PowerCalRatio1(i, 2) = 0.5 * (PowerCalRatio1(i, 2) + P_LSNA.Re)
                PowerCalRatio1(i, 3) = 0.5 * (PowerCalRatio1(i, 3) + P_LSNA.Im)

            Next i

        End If

        'Get rid of the phase, which is not relvant here
        For i As Integer = 1 To MechanismList1.FrequencyList.NRows
            PowerCalRatio1(i, 2) = Math.Sqrt(PowerCalRatio1(i, 2) * PowerCalRatio1(i, 2) + PowerCalRatio1(i, 3) * PowerCalRatio1(i, 3))
            PowerCalRatio1(i, 3) = 0.0
            'If PM_Port_SelectedIndex = 1 Then PowerCalRatio1(i, 2) = 1.0 / PowerCalRatio1(i, 2) 'invert the result if the power meter is on port 2
        Next i

        Return PowerCalRatio1

    End Function


    ''' <summary>
    ''' Calculate the ratio of the phase the LSNA measures to the actual phase of the signal incident on it
    ''' </summary>
    ''' <param name="MechanismList1"></param>
    ''' <param name="HPR_Port_SelectedIndex">The indes of the combo box selecting the port that the HPR is connected to. 0=P1, 1=P2</param>
    ''' <returns>The ratio of the power the LSNA thinks that it generated to the power that was actually generated</returns>
    ''' <remarks>We are assuming that the port numbers did not get changed in a multiport calibration.</remarks>
    Public Function PhaseCalRatio(ByVal MechanismList1 As MechanismList, ByVal HPR_Port_SelectedIndex As Integer) As RealMatrix

        'Find the power calibration ratio for the before calibration
        Dim PhaseCalRatio1 As New RealMatrix(MechanismList1.FrequencyList.NRows, 3, ".complex")
        PhaseCalRatio1.Vector(1) = MechanismList1.FrequencyList

        'Pull out the calibration matrix from MechanismList1.PreSolution and get the port-1 s-parameter matrix.
        Dim SParCalibrationPort1 As New RealMatrix(MechanismList1.FrequencyList.NRows, 9, ".s2p")
        SParCalibrationPort1.Fill(GetPortCalibration(MechanismList1, HPR_Port_SelectedIndex))

        'Find the scattering parameters of the adapter and mismatches
        Dim SPar As New RealMatrix(MechanismList1.FrequencyList.NRows, 9, ".s2p")
        SPar.InitializeAsSParams()
        SPar.Vector(1) = MechanismList1.FrequencyList
        Dim SPar1 As New RealMatrix(MechanismList1.FrequencyList.NRows, 9, ".s2p")
        SPar1.InitializeAsSParams()
        SPar1.Vector(1) = MechanismList1.FrequencyList
        Dim SPar2 As New RealMatrix(MechanismList1.FrequencyList.NRows, 9, ".s2p")
        SPar2.InitializeAsSParams()
        SPar2.Vector(1) = MechanismList1.FrequencyList
        SPar1 = myBeforeHPRAdapterS2P.getSParams(MechanismList1)
        SPar = myBeforeHPRMismatchS2P.getSParams(MechanismList1)
        If SPar(1, 4) = 0 And SPar(1, 5) = 0 And SPar(1, 6) = 0 And SPar(1, 7) = 0 Then 'User probably did not realize that transmission should be set to 1 
            For i As Integer = 1 To SPar.NRows
                SPar(i, 4) = 1.0 : SPar(i, 6) = 1.0
            Next i
        End If
        SPar2 = SPar1.CascadeSParameters(SPar)
        SPar = SParCalibrationPort1.CascadeSParameters(SPar2)

        'The signal a1 going into the LSNA and the signal generated by the HPR read by the power meter.
        Dim P_LSNA1 As New RealMatrix(MechanismList1.FrequencyList.NRows, 5, ".w1p")
        Dim P_HPR1 As New RealMatrix(MechanismList1.FrequencyList.NRows, 5, ".w1p")
        P_LSNA1 = myBeforeHPRMeasurementW1P.getSParams(MechanismList1)
        P_HPR1 = myBeforeHPRW1P.getSParams(MechanismList1)

        'Calculate the signal ratio for the before calibration
        For i As Integer = 1 To MechanismList1.FrequencyList.NRows

            'The signal b1 coming out of the HPR.
            Dim P_HPR As New Complex(P_HPR1(i, 4), P_HPR1(i, 5))

            'Calculate the actual signal b1=S12*a2 going into the LSNA.
            P_HPR = P_HPR * toComplex(SPar(i, 6), SPar(i, 7))

            'Compare the signal going into the LSNA with the b1 coming from the HPR and form the ratio.
            P_HPR = toComplex(P_LSNA1(i, 4), P_LSNA1(i, 5)) / P_HPR

            'Stuff into the output matrix
            PhaseCalRatio1(i, 2) = P_HPR.Re : PhaseCalRatio1(i, 3) = P_HPR.Im

        Next i

        'Find the power calibration ratio for the after calibration and average in the answers.
        If MechanismList1.BeforeAfterSwitch >= 0 And Not (myAfterHPRMeasurementW1P Is Nothing) And Not (myAfterHPRW1P Is Nothing) Then

            'Find the scattering parameters of the adapter and mismatches
            SPar.InitializeAsSParams()
            SPar1.InitializeAsSParams()
            SPar2.InitializeAsSParams()
            SPar1 = myAfterHPRAdapterS2P.getSParams(MechanismList1)
            SPar = myAfterHPRMismatchS2P.getSParams(MechanismList1)
            If SPar(1, 4) = 0 And SPar(1, 5) = 0 And SPar(1, 6) = 0 And SPar(1, 7) = 0 Then 'User probably did not realize that transmission should be set to 1 
                For i As Integer = 1 To SPar.NRows
                    SPar(i, 4) = 1.0 : SPar(i, 6) = 1.0
                Next i
            End If
            SPar2 = SPar1.CascadeSParameters(SPar)
            SPar = SParCalibrationPort1.CascadeSParameters(SPar2)

            'The signal a1 going into the LSNA and the signal generated by the HPR read by the power meter.
            P_LSNA1 = myAfterHPRMeasurementW1P.getSParams(MechanismList1)
            P_HPR1 = myAfterHPRW1P.getSParams(MechanismList1)

            'Calculate the signal ratio for the before calibration
            For i As Integer = 1 To MechanismList1.FrequencyList.NRows

                'The signal b1 coming out of the HPR.
                Dim P_HPR As New Complex(P_HPR1(i, 4), P_HPR1(i, 5))

                'Calculate the signal b1 going into the LSNA.
                P_HPR = P_HPR * toComplex(SPar(i, 6), SPar(i, 7))

                'Compare the signal going into the LSNA with the a1 from the HPR and form the ratio.
                P_HPR = toComplex(P_LSNA1(i, 4), P_LSNA1(i, 5)) / P_HPR

                'Stuff into the output matrix, averaging the two results.
                PhaseCalRatio1(i, 2) = 0.5 * (PhaseCalRatio1(i, 2) + P_HPR.Re)
                PhaseCalRatio1(i, 3) = 0.5 * (PhaseCalRatio1(i, 3) + P_HPR.Im)

            Next i

        End If

        'Get rid of the amplitude, which is not relvant here
        For i As Integer = 1 To MechanismList1.FrequencyList.NRows
            Dim Ratio As New Complex(PhaseCalRatio1(i, 2), PhaseCalRatio1(i, 3))
            Ratio = Ratio / Complex_Number_Class.Abs(Ratio)
            PhaseCalRatio1(i, 2) = Ratio.Re
            PhaseCalRatio1(i, 3) = Ratio.Im
            'If HPR_Port_SelectedIndex = 1 Then PhaseCalRatio1(i, 3) = -Ratio.Im 'invert the result if the HPR was on port 2
        Next i

        Return PhaseCalRatio1

    End Function

    ''' <summary>
    ''' Create a list of all of the before and after calibration models
    ''' </summary>
    ''' <param name="MechanismList1">The mechanism list.</param>
    ''' <param name="Port1Connection">Only standards connected to these ports will be enabled.</param>
    ''' <param name="Port2Connection">Only standards connected to these ports will be enabled.</param>
    ''' <remarks>Setting Port1Connection to 0 enables all standards.</remarks>
    Public Sub CollapseBeforeAndAfter(ByVal MechanismList1 As MechanismList, ByVal Port1Connection As Integer, ByVal Port2Connection As Integer)

        'Set the current port connections
        myCurrentPort1Connection = Port1Connection
        myCurrentPort2Connection = Port2Connection

        'Use the before/after calibration status to assemble the correct list of standards.
        Dim FullCount As Integer = myBeforeCalibrationModels.Count + myAfterCalibrationModels.Count, modIndex As Integer = -1
        ReDim myCalibrationModels(FullCount - 1)
        ReDim myCalibrationModelNames(FullCount - 1)
        ReDim myIgnoreCalibrationModels(FullCount - 1)
        ReDim myCalibrationStdTypes(FullCount - 1)
        ReDim myCalibrationLengths(FullCount - 1)
        ReDim myCalibrationMeasurements(FullCount - 1)
        ReDim myPort1Connection(FullCount - 1)
        ReDim myPort2Connection(FullCount - 1)
        If myBeforeCalibrationModels.Count > 0 Then
            For k As Integer = 0 To myBeforeCalibrationModels.Count - 1
                modIndex = modIndex + 1
                myCalibrationModels(modIndex) = myBeforeCalibrationModels(k)
                myCalibrationModelNames(modIndex) = myBeforeCalibrationModelNames(k)
                myCalibrationStdTypes(modIndex) = myBeforeCalibrationStdTypes(k)
                myCalibrationLengths(modIndex) = myBeforeCalibrationLengths(k)
                myCalibrationMeasurements(modIndex) = myBeforeCalibrationMeasurements(k)
                myPort1Connection(modIndex) = myBeforePort1Connection(k)
                myPort2Connection(modIndex) = myBeforePort2Connection(k)
                myIgnoreCalibrationModels(modIndex) = (MechanismList1.BeforeAfterSwitch > 0)
                'If Port1Connection > 0 And Port1Connection <> myPort1Connection(modIndex) Then myIgnoreCalibrationModels(modIndex) = True
                'If Port2Connection > 0 And Port2Connection <> myPort2Connection(modIndex) Then myIgnoreCalibrationModels(modIndex) = True
            Next k
        End If
        If myAfterCalibrationModels.Count > 0 Then
            For k As Integer = 0 To myAfterCalibrationModels.Count - 1
                modIndex = modIndex + 1
                myCalibrationModels(modIndex) = myAfterCalibrationModels(k)
                myCalibrationModelNames(modIndex) = myAfterCalibrationModelNames(k)
                myCalibrationStdTypes(modIndex) = myAfterCalibrationStdTypes(k)
                myCalibrationLengths(modIndex) = myAfterCalibrationLengths(k)
                myCalibrationMeasurements(modIndex) = myAfterCalibrationMeasurements(k)
                myPort1Connection(modIndex) = myAfterPort1Connection(k)
                myPort2Connection(modIndex) = myAfterPort2Connection(k)
                myIgnoreCalibrationModels(modIndex) = (MechanismList1.BeforeAfterSwitch < 0)
                'If Port1Connection > 0 And Port1Connection <> myPort1Connection(modIndex) Then myIgnoreCalibrationModels(modIndex) = True
                'If Port2Connection > 0 And Port2Connection <> myPort2Connection(modIndex) Then myIgnoreCalibrationModels(modIndex) = True
            Next k
        End If

        'Get the SwitchTerms and isolation terms.
        Dim NSwitch As Integer = 0, NIso As Integer = 0
        Dim Switch As New RealMatrix(MechanismList1.FrequencyList.NRows, 9, ".switch")
        Dim Iso As New RealMatrix(MechanismList1.FrequencyList.NRows, 9, ".iso")
        For k As Integer = 0 To FullCount - 1
            'Get the switch terms
            If Not myIgnoreCalibrationModels(k) And InStr(myCalibrationStdTypes(k), "Switch term") > 0 Then
                Dim Switch1 As RealMatrix = myCalibrationMeasurements(k).getSParams(MechanismList1)
                Switch = Switch + Switch1 : NSwitch = NSwitch + 1
            End If
            'Get the isolation terms
            If Not myIgnoreCalibrationModels(k) And InStr(myCalibrationStdTypes(k), "Isolation standard") > 0 Then
                Dim Iso1 As RealMatrix = myCalibrationMeasurements(k).getSParams(MechanismList1)
                Iso = Iso + Iso1 : NIso = NIso + 1
            End If
        Next k
        'Get mean of various switch and isolation terms
        If NSwitch > 1 Then
            For k As Integer = 1 To MechanismList1.FrequencyList.NRows
                Switch(k, 2) = Switch(k, 2) / CDbl(NSwitch) : Switch(k, 3) = Switch(k, 3) / CDbl(NSwitch)
                Switch(k, 4) = Switch(k, 4) / CDbl(NSwitch) : Switch(k, 5) = Switch(k, 5) / CDbl(NSwitch)
            Next k
        End If
        If NIso > 1 Then
            For k As Integer = 1 To MechanismList1.FrequencyList.NRows
                Iso(k, 2) = Iso(k, 2) / CDbl(NIso) : Iso(k, 3) = Iso(k, 3) / CDbl(NIso)
                Iso(k, 4) = Iso(k, 4) / CDbl(NIso) : Iso(k, 5) = Iso(k, 5) / CDbl(NIso)
                Iso(k, 6) = Iso(k, 6) / CDbl(NIso) : Iso(k, 7) = Iso(k, 7) / CDbl(NIso)
                Iso(k, 8) = Iso(k, 8) / CDbl(NIso) : Iso(k, 9) = Iso(k, 9) / CDbl(NIso)
            Next k
        End If
        Switch.Vector(1) = MechanismList1.FrequencyList
        MechanismList1.SwitchTerms = Switch
        Iso.Vector(1) = MechanismList1.FrequencyList
        MechanismList1.IsolationTerms = Iso

    End Sub


    ''' <summary>
    ''' Create a list of all of the .snp before and after calibration models created from the .wnp data
    ''' </summary>
    ''' <param name="MechanismList1">The mechanism list.</param>
    ''' <param name="Port1Connection">Only standards connected to these ports will be enabled.</param>
    ''' <param name="Port2Connection">Only standards connected to these ports will be enabled.</param>
    ''' <remarks>The .snp measurements are switch-term corrected by this subroutine.
    ''' Setting Port1Connection to 0 enables all standards.
    ''' Set Port2Connection to 0 for one-port calibrations.
    ''' Switch terms are automatically put into the Mechanism List for later use if need be for these port choices.</remarks>
    Public Sub CollapseBeforeAndAfterWnpToSnp(ByVal MechanismList1 As MechanismList, ByVal Port1Connection As Integer, ByVal Port2Connection As Integer)

        'Set the current port connections
        If Port1Connection > 0 Then myCurrentPort1Connection = Port1Connection
        If Port2Connection > 0 Then myCurrentPort2Connection = Port2Connection

        'Use the before/after calibration status to assemble the correct list of standards.
        Dim FullCount As Integer = myBeforeCalibrationModels.Count + myAfterCalibrationModels.Count, modIndex As Integer = -1
        ReDim myCalibrationModels(FullCount - 1)
        ReDim myCalibrationModelNames(FullCount - 1)
        ReDim myIgnoreCalibrationModels(FullCount - 1)
        ReDim myCalibrationStdTypes(FullCount - 1)
        ReDim myCalibrationLengths(FullCount - 1)
        ReDim myCalibrationMeasurements(FullCount - 1)
        ReDim myPort1Connection(FullCount - 1)
        ReDim myPort2Connection(FullCount - 1)

        'Other useful matricies.

        'The port 1 and port 2 driven waves
        Dim myWaveRepresentation As New RealMatrix(MechanismList1.FrequencyList.NRows, 1 + 4 * NumberOfPorts * NumberOfPorts, ".w" + NumberOfPorts.ToString + "p")

        'A place to collect the switch terms.
        Dim NSwitch As Integer = 0
        Dim Switch As New RealMatrix(MechanismList1.FrequencyList.NRows, 9, ".switch")

        'A quick check of the port connections. This should never happen.
        If Port1Connection > 0 And Port2Connection > 0 And Port1Connection = Port2Connection Then
            Throw New ApplicationException("Calibration Engine:CollapseBeforeAndAfterWnpToSnp: Port1Connection = Port2Connection")
        End If

        'The before measurements.
        If myBeforeCalibrationModels.Count > 0 Then
            For k As Integer = 0 To myBeforeCalibrationModels.Count - 1
                modIndex = modIndex + 1
                myCalibrationModels(modIndex) = myBeforeCalibrationModels(k)
                myCalibrationModelNames(modIndex) = myBeforeCalibrationModelNames(k)
                myCalibrationStdTypes(modIndex) = myBeforeCalibrationStdTypes(k)
                myCalibrationLengths(modIndex) = myBeforeCalibrationLengths(k)
                myPort1Connection(modIndex) = myBeforePort1Connection(k)
                myPort2Connection(modIndex) = myBeforePort2Connection(k)
                'Use BeforeAfterSwitch to ignore some standards
                myIgnoreCalibrationModels(modIndex) = (MechanismList1.BeforeAfterSwitch > 0)
                'Ignore standards that don't have the correct port connections.
                If Port1Connection > 0 And Port1Connection <> myPort1Connection(modIndex) Then myIgnoreCalibrationModels(modIndex) = True
                If Port2Connection > 0 And Port2Connection <> myPort2Connection(modIndex) Then myIgnoreCalibrationModels(modIndex) = True
                'Ignore reciprocal standards if this is a one-port calibration.
                If Port2Connection = 0 And InStr(myCalibrationStdTypes(modIndex), "Reciprocal") > 0 Then myIgnoreCalibrationModels(modIndex) = True

                'Convert the .wnp files to .snp files.
                If Not myIgnoreCalibrationModels(modIndex) Then
                    'Interpolate the measurements ans stick them in myWavePXRepresentation
                    myWaveRepresentation = myBeforeCalibrationMeasurements(k).getSParams(MechanismList1)
                    Dim myRealMatrixContainer As New RealMatrixContainer(ConvertWnpToSnp(myWaveRepresentation, Port1Connection, Port2Connection, MechanismList1, False))
                    myCalibrationMeasurements(modIndex) = myRealMatrixContainer
                    'If this is a Reciprocal, Thru, or Line, get the switch terms
                    If Port1Connection > 0 And Port2Connection > 0 And InStr(myCalibrationStdTypes(modIndex), "Reciprocal") + InStr(myCalibrationStdTypes(modIndex), "Thru") + InStr(myCalibrationStdTypes(modIndex), "Line") + InStr(myCalibrationStdTypes(modIndex), "Switch term") > 0 Then
                        Dim Switch1 As RealMatrix = ConvertWnpToSnp(myWaveRepresentation, Port1Connection, Port2Connection, MechanismList1, True)
                        Switch = Switch + Switch1 : NSwitch = NSwitch + 1
                    End If
                End If

            Next k
        End If

        'The after measurements.
        If myAfterCalibrationModels.Count > 0 Then
            For k As Integer = 0 To myAfterCalibrationModels.Count - 1
                modIndex = modIndex + 1
                myCalibrationModels(modIndex) = myAfterCalibrationModels(k)
                myCalibrationModelNames(modIndex) = myAfterCalibrationModelNames(k)
                myCalibrationStdTypes(modIndex) = myAfterCalibrationStdTypes(k)
                myCalibrationLengths(modIndex) = myAfterCalibrationLengths(k)
                myPort1Connection(modIndex) = myAfterPort1Connection(k)
                myPort2Connection(modIndex) = myAfterPort2Connection(k)
                'Use BeforeAfterSwitch to ignore some standards
                myIgnoreCalibrationModels(modIndex) = (MechanismList1.BeforeAfterSwitch < 0)
                'Ignore standards that don't have the correct port connections.
                If Port1Connection > 0 And Port1Connection <> myPort1Connection(modIndex) Then myIgnoreCalibrationModels(modIndex) = True
                If Port2Connection > 0 And Port2Connection <> myPort2Connection(modIndex) Then myIgnoreCalibrationModels(modIndex) = True
                'Ignore reciprocal standards if this is a one-port calibration.
                If Port2Connection = 0 And InStr(myCalibrationStdTypes(modIndex), "Reciprocal") > 0 Then myIgnoreCalibrationModels(modIndex) = True

                'Convert the .wnp files to .snp files.
                If Not myIgnoreCalibrationModels(modIndex) Then
                    'Convert the wave measurements and stick them in myWavePXRepresentation
                    myWaveRepresentation = myAfterCalibrationMeasurements(k).getSParams(MechanismList1)
                    Dim myRealMatrixContainer As New RealMatrixContainer(ConvertWnpToSnp(myWaveRepresentation, Port1Connection, Port2Connection, MechanismList1, False))
                    myCalibrationMeasurements(modIndex) = myRealMatrixContainer
                    'If this is a Reciprocal, Thru, or Line, get the switch terms
                    If Port1Connection > 0 And Port2Connection > 0 And InStr(myCalibrationStdTypes(modIndex), "Reciprocal") + InStr(myCalibrationStdTypes(modIndex), "Thru") + InStr(myCalibrationStdTypes(modIndex), "Line") + InStr(myCalibrationStdTypes(modIndex), "Switch term") > 0 Then
                        Dim Switch1 As RealMatrix = ConvertWnpToSnp(myWaveRepresentation, Port1Connection, Port2Connection, MechanismList1, True)
                        Switch = Switch + Switch1 : NSwitch = NSwitch + 1
                    End If
                End If

            Next k
        End If

        'If we have two-ports, we can try to set switch terms and switch-term correct the files.
        If Port1Connection > 0 And Port2Connection > 0 Then

            'Get mean of switch terms
            If NSwitch > 1 Then
                For k As Integer = 1 To MechanismList1.FrequencyList.NRows
                    Switch(k, 2) = Switch(k, 2) / CDbl(NSwitch) : Switch(k, 3) = Switch(k, 3) / CDbl(NSwitch)
                    Switch(k, 4) = Switch(k, 4) / CDbl(NSwitch) : Switch(k, 5) = Switch(k, 5) / CDbl(NSwitch)
                Next k
            End If

            'Put the SwitchTerms into the Mechanism List
            Switch.Vector(1) = MechanismList1.FrequencyList
            MechanismList1.SwitchTerms = Switch

            ''Switch-term correct the measurements.
            ''This does not work, as the calibration engines later expect to correct for the switch terms in MechanismList1
            'For k As Integer = 0 To FullCount - 1

            '    'Perform the switch term corrections
            '    Dim myCalibrationMeasurementsSwitchCorrected As New RealMatrix(MechanismList1.FrequencyList.NRows, 9)
            '    myCalibrationMeasurementsSwitchCorrected = myCalibrationMeasurementsNoSwitch(k).SwitchTermCorrection(myCalibrationMeasurementsNoSwitch(k), MechanismList1)

            '    'Save the switch-term corrected measurements away.
            '    Dim myRealMatrixContainer As New RealMatrixContainer(myCalibrationMeasurementsSwitchCorrected)
            '    myCalibrationMeasurements(k) = myRealMatrixContainer

            'Next k

        End If

    End Sub

    ''' <summary>
    ''' Create a list of all of the .snp before and after calibration models created from the .wnp data
    ''' </summary>
    ''' <param name="MechanismList1">The mechanism list.</param>
    ''' <remarks>The .snp measurements are switch-term corrected by this subroutine.
    ''' Setting Port1Connection to 0 enables all standards.
    ''' Set Port2Connection to 0 for one-port calibrations.
    ''' Switch terms are automatically put into the Mechanism List for later use if need be for these port choices.</remarks>
    Public Sub CollapseBeforeAndAfterWnp(ByVal MechanismList1 As MechanismList)

        'Use the before/after calibration status to assemble the correct list of standards.
        Dim FullCount As Integer = myBeforeCalibrationModels.Count + myAfterCalibrationModels.Count, modIndex As Integer = -1
        ReDim myCalibrationModels(FullCount - 1)
        ReDim myCalibrationModelNames(FullCount - 1)
        ReDim myIgnoreCalibrationModels(FullCount - 1)
        ReDim myCalibrationStdTypes(FullCount - 1)
        ReDim myCalibrationLengths(FullCount - 1)
        ReDim myCalibrationMeasurements(FullCount - 1)
        ReDim myPort1Connection(FullCount - 1)
        ReDim myPort2Connection(FullCount - 1)

        'Other useful matricies.

        'The port 1 and port 2 driven waves
        Dim myWaveRepresentation As New RealMatrix(MechanismList1.FrequencyList.NRows, 1 + 4 * NumberOfPorts * NumberOfPorts, ".w" + NumberOfPorts.ToString + "p")

        'A place to collect the switch terms.
        Dim NSwitch(NumberOfPorts) As Integer
        Dim Switch As New RealMatrix(MechanismList1.FrequencyList.NRows, 1 + 2 * NumberOfPorts, ".switch")

        'The before measurements.
        If myBeforeCalibrationModels.Count > 0 Then
            For k As Integer = 0 To myBeforeCalibrationModels.Count - 1
                modIndex = modIndex + 1
                myCalibrationModels(modIndex) = myBeforeCalibrationModels(k)
                myCalibrationModelNames(modIndex) = myBeforeCalibrationModelNames(k)
                myCalibrationStdTypes(modIndex) = myBeforeCalibrationStdTypes(k)
                myCalibrationLengths(modIndex) = myBeforeCalibrationLengths(k)
                myPort1Connection(modIndex) = myBeforePort1Connection(k)
                myPort2Connection(modIndex) = myBeforePort2Connection(k)
                'Use BeforeAfterSwitch to ignore some standards
                myIgnoreCalibrationModels(modIndex) = (MechanismList1.BeforeAfterSwitch > 0)
                'Ignore standards that don't have the correct port connections.
                If myPort1Connection(modIndex) = 0 Then myIgnoreCalibrationModels(modIndex) = True
                If myPort2Connection(modIndex) = 0 Then myIgnoreCalibrationModels(modIndex) = True

                'Convert the .wnp files to .snp files.
                If Not myIgnoreCalibrationModels(modIndex) Then

                    'Interpolate the measurements ans stick them in myWavePXRepresentation
                    myWaveRepresentation = myBeforeCalibrationMeasurements(k).getSParams(MechanismList1)

                    'Return the whole .wnp file.
                    myCalibrationMeasurements(modIndex) = myWaveRepresentation

                    'If this is a Reciprocal, Thru, or Line, get the switch terms
                    'Note that ConvertWnpToSnp now returns zero switch terms. We could really get rid of this if we wanted. But it does no harm.
                    If InStr(myCalibrationStdTypes(modIndex), "Reciprocal") + InStr(myCalibrationStdTypes(modIndex), "Thru") + InStr(myCalibrationStdTypes(modIndex), "Line") + InStr(myCalibrationStdTypes(modIndex), "Switch term") > 0 Then
                        Dim Switch1 As RealMatrix = ConvertWnpToSnp(myWaveRepresentation, myPort1Connection(modIndex), myPort2Connection(modIndex), MechanismList1, True)
                        Switch.Vector(2 * myPort1Connection(modIndex)) = Switch.Vector(2 * myPort1Connection(modIndex)) + Switch1.Vector(2)
                        Switch.Vector(2 * myPort1Connection(modIndex) + 1) = Switch.Vector(2 * myPort1Connection(modIndex) + 1) + Switch1.Vector(3)
                        Switch.Vector(2 * myPort2Connection(modIndex)) = Switch.Vector(2 * myPort2Connection(modIndex)) + Switch1.Vector(4)
                        Switch.Vector(2 * myPort2Connection(modIndex) + 1) = Switch.Vector(2 * myPort2Connection(modIndex) + 1) + Switch1.Vector(5)
                        NSwitch(myPort1Connection(modIndex)) = NSwitch(myPort1Connection(modIndex)) + 1
                        NSwitch(myPort2Connection(modIndex)) = NSwitch(myPort2Connection(modIndex)) + 1
                    End If
                End If

            Next k
        End If


        'The After measurements.
        If myAfterCalibrationModels.Count > 0 Then
            For k As Integer = 0 To myAfterCalibrationModels.Count - 1
                modIndex = modIndex + 1
                myCalibrationModels(modIndex) = myAfterCalibrationModels(k)
                myCalibrationModelNames(modIndex) = myAfterCalibrationModelNames(k)
                myCalibrationStdTypes(modIndex) = myAfterCalibrationStdTypes(k)
                myCalibrationLengths(modIndex) = myAfterCalibrationLengths(k)
                myPort1Connection(modIndex) = myAfterPort1Connection(k)
                myPort2Connection(modIndex) = myAfterPort2Connection(k)
                'Use BeforeAfterSwitch to ignore some standards
                myIgnoreCalibrationModels(modIndex) = (MechanismList1.BeforeAfterSwitch > 0)
                'Ignore standards that don't have the correct port connections.
                If myPort1Connection(modIndex) = 0 Then myIgnoreCalibrationModels(modIndex) = True
                If myPort2Connection(modIndex) = 0 Then myIgnoreCalibrationModels(modIndex) = True

                'Convert the .wnp files to .snp files.
                If Not myIgnoreCalibrationModels(modIndex) Then
                    'Interpolate the measurements ans stick them in myWavePXRepresentation
                    myWaveRepresentation = myAfterCalibrationMeasurements(k).getSParams(MechanismList1)

                    'Return the whole .wnp file.
                    myCalibrationMeasurements(modIndex) = myWaveRepresentation

                    'If this is a Reciprocal, Thru, or Line, get the switch terms
                    If InStr(myCalibrationStdTypes(modIndex), "Reciprocal") + InStr(myCalibrationStdTypes(modIndex), "Thru") + InStr(myCalibrationStdTypes(modIndex), "Line") + InStr(myCalibrationStdTypes(modIndex), "Switch term") > 0 Then
                        Dim Switch1 As RealMatrix = ConvertWnpToSnp(myWaveRepresentation, myPort1Connection(modIndex), myPort2Connection(modIndex), MechanismList1, True)
                        Switch.Vector(2 * myPort1Connection(modIndex)) = Switch.Vector(2 * myPort1Connection(modIndex)) + Switch1.Vector(2)
                        Switch.Vector(2 * myPort1Connection(modIndex) + 1) = Switch.Vector(2 * myPort1Connection(modIndex) + 1) + Switch1.Vector(3)
                        Switch.Vector(2 * myPort2Connection(modIndex)) = Switch.Vector(2 * myPort2Connection(modIndex)) + Switch1.Vector(4)
                        Switch.Vector(2 * myPort2Connection(modIndex) + 1) = Switch.Vector(2 * myPort2Connection(modIndex) + 1) + Switch1.Vector(5)
                        NSwitch(myPort1Connection(modIndex)) = NSwitch(myPort1Connection(modIndex)) + 1
                        NSwitch(myPort2Connection(modIndex)) = NSwitch(myPort2Connection(modIndex)) + 1
                    End If
                End If

            Next k
        End If

        'Try to set switch terms and switch-term correct the files.

        'Get mean of switch terms
        For ISwitch = 1 To NumberOfPorts
            If NSwitch(ISwitch) > 1 Then
                For k As Integer = 1 To MechanismList1.FrequencyList.NRows
                    Switch(k, 2 * ISwitch) = Switch(k, 2 * ISwitch) / CDbl(NSwitch(ISwitch))
                    Switch(k, 2 * ISwitch + 1) = Switch(k, 2 * ISwitch + 1) / CDbl(NSwitch(ISwitch))
                Next k
            End If
        Next ISwitch

        'Put the SwitchTerms into the Mechanism List
        Switch.Vector(1) = MechanismList1.FrequencyList
        MechanismList1.SwitchTerms = Switch

    End Sub


    ''' <summary>
    ''' Returns a list of the reciprocal measurements and models for this port setting.
    ''' </summary>
    ''' <param name="MechanismList1">The mechanism list.</param>
    ''' <param name="NReciprocals">The number of reciprocals for this port setting.</param>
    ''' <param name="ReciprocalMeasurements"></param>
    ''' <param name="ReciprocalModels"></param>
    ''' <remarks>The number of reciprocals for this port setting must be greater than 0 or dummy data is returned.</remarks>
    Public Sub ReturnCurrentReciprocalMeasurements(ByVal MechanismList1 As MechanismList, ByRef NReciprocals As Integer, ByRef ReciprocalMeasurements As RealMatrix(), ByRef ReciprocalModels As RealMatrix())

        'Return a list of the reciprocals for this port setting.
        Dim NReciprocal As Integer = -1
        ReDim ReciprocalMeasurements(0), ReciprocalModels(0)
        For k As Integer = 0 To myCalibrationMeasurements.Count - 1
            If Not myIgnoreCalibrationModels(k) Then
                If InStr(myCalibrationStdTypes(k), "Reciprocal") + InStr(myCalibrationStdTypes(k), "Thru") + InStr(myCalibrationStdTypes(k), "Line") + InStr(myCalibrationStdTypes(k), "Switch term") > 0 Then
                    NReciprocal = NReciprocal + 1
                    ReDim Preserve ReciprocalMeasurements(NReciprocal)
                    ReDim Preserve ReciprocalModels(NReciprocal)
                    ReciprocalMeasurements(NReciprocal) = myCalibrationMeasurements(k).getSParams(MechanismList1)
                    ReciprocalModels(NReciprocal) = myCalibrationModels(k).getSParams(MechanismList1)
                End If
            End If
        Next

    End Sub

    ''' <summary>
    ''' Convert the .wnp wave measurements to .s2p measurements for the two ports between which we are calibrating
    ''' </summary>
    ''' <param name="myWaveRepresentation">The input .wnp wave measurement.</param>
    ''' <param name="Port1Connection">The first port we are calibrating at.</param>
    ''' <param name="Port2Connection">The second port we are calibrating at.</param>
    ''' <param name="MechanismList1"></param>
    ''' <param name="IsSwitchTerms">In case we have a request for switch terms. Just set them to zero, as wave measurements don't need them.</param>
    ''' <returns>The two-port scattering parameters we can use to calibrate with.</returns>
    ''' <remarks>This replaces the old code that tries to create the regular VNA model with switch terms with new code that applies an LSNA model with no switch terms.</remarks>
    Public Shared Function ConvertWnpToSnp(ByVal myWaveRepresentation As RealMatrix, ByVal Port1Connection As Integer, ByVal Port2Connection As Integer, ByVal MechanismList1 As MechanismList, Optional ByVal IsSwitchTerms As Boolean = False) As RealMatrix

        'Set up the variables we will need.
        Dim NPorts As Integer = myWaveRepresentation.NPorts
        Dim mySParamRepresentation As New RealMatrix(MechanismList1.FrequencyList.NRows, 9, ".s2p")
        If NPorts = 0 Then Throw New ApplicationException("CalibrationEngine.ConvertWnpToSnp: NPorts=0")

        'Do this as if we are using an LSNA. No need for switch terms here.

        If Not IsSwitchTerms Then

            If Port2Connection > 0 And NPorts <> 1 Then 'If this is not a one-port, return two-port parameters

                'Get only the two-ports that matter here, Port1Connection and Port2Connection, out of the .wnp file.
                Dim myWaveRepresentationSubMatrix As New RealMatrix(MechanismList1.FrequencyList.NRows, 17, ".w2p")
                For k As Integer = 1 To 4
                    'Get a1 and b1 with power at port 1
                    myWaveRepresentationSubMatrix.Vector(1 + k) = myWaveRepresentation.Vector(1 + k + 4 * (Port1Connection - 1) + 4 * NPorts * (Port1Connection - 1))
                    'Get a2 and b2 with power at port 1
                    myWaveRepresentationSubMatrix.Vector(5 + k) = myWaveRepresentation.Vector(1 + k + 4 * (Port2Connection - 1) + 4 * NPorts * (Port1Connection - 1))
                    'Get a1 and b1 with power at port 2
                    myWaveRepresentationSubMatrix.Vector(9 + k) = myWaveRepresentation.Vector(1 + k + 4 * (Port1Connection - 1) + 4 * NPorts * (Port2Connection - 1))
                    'Get a2 and b2 with power at port 2
                    myWaveRepresentationSubMatrix.Vector(13 + k) = myWaveRepresentation.Vector(1 + k + 4 * (Port2Connection - 1) + 4 * NPorts * (Port2Connection - 1))
                Next k
                'Convert the .w2p file to a .s2p file.
                mySParamRepresentation = myWaveRepresentationSubMatrix.WnP_to_SnP  'CustomFormControls.WnP_to_SnP(mySParamRepresentation)

            Else        'This is a one-port calibration. Return just S11

                mySParamRepresentation.Vector(1) = MechanismList1.FrequencyList   'Frequency
                Dim a As New ComplexMatrix(MechanismList1.FrequencyList.NRows), b As New ComplexMatrix(MechanismList1.FrequencyList.NRows)
                Dim S As New Complex()
                'SF11 = b1m/a1m, drive on Port1Connection
                a = myWaveRepresentation.Vector(2 + 4 * (Port1Connection - 1) + 4 * NPorts * (Port1Connection - 1)) + toComplex(0.0, 1.0) * myWaveRepresentation.Vector(3 + 4 * (Port1Connection - 1) + 4 * NPorts * (Port1Connection - 1))
                b = myWaveRepresentation.Vector(4 + 4 * (Port1Connection - 1) + 4 * NPorts * (Port1Connection - 1)) + toComplex(0.0, 1.0) * myWaveRepresentation.Vector(5 + 4 * (Port1Connection - 1) + 4 * NPorts * (Port1Connection - 1))
                For i1 As Integer = 1 To MechanismList1.FrequencyList.NRows
                    S = b(i1) / a(i1)
                    mySParamRepresentation(i1, 2) = S.Re : mySParamRepresentation(i1, 3) = S.Im
                Next i1

            End If

        Else    'Don't do anything. Just let the switch terms be zero.

        End If



        ''This was the old code I was using. This tries to create the regular VNA model with switch terms.
        ''This seems completly wrong. Should use the approach taken in the post processor that does this.
        ''What we need is an LSNA-style wave-parameter to S-parameter conversion that is independent of the switch terms. 
        ''This is a regular scattering-parameter measurement. Form it from the wave parameters.

        ''Do this as if we were using a regular VNA approach with switch terms
        'mySParamRepresentation.Vector(1) = MechanismList1.FrequencyList   'Frequency
        'Dim a As New ComplexMatrix(MechanismList1.FrequencyList.NRows), b As New ComplexMatrix(MechanismList1.FrequencyList.NRows)
        'Dim S As New Complex()

        'If Not IsSwitchTerms Then

        '    'SF11 = b1m/a1m, drive on Port1Connection
        '    a = myWaveRepresentation.Vector(2 + 4 * (Port1Connection - 1) + 4 * NPorts * (Port1Connection - 1)) + toComplex(0.0, 1.0) * myWaveRepresentation.Vector(3 + 4 * (Port1Connection - 1) + 4 * NPorts * (Port1Connection - 1))
        '    b = myWaveRepresentation.Vector(4 + 4 * (Port1Connection - 1) + 4 * NPorts * (Port1Connection - 1)) + toComplex(0.0, 1.0) * myWaveRepresentation.Vector(5 + 4 * (Port1Connection - 1) + 4 * NPorts * (Port1Connection - 1))
        '    For i1 As Integer = 1 To MechanismList1.FrequencyList.NRows
        '        S = b(i1) / a(i1)
        '        mySParamRepresentation(i1, 2) = S.Re : mySParamRepresentation(i1, 3) = S.Im
        '    Next i1
        '    If Port2Connection > 0 Then 'If this is not a one-port
        '        'SF21 = b2m/a1m, drive on Port1Connection
        '        b = myWaveRepresentation.Vector(4 + 4 * (Port2Connection - 1) + 4 * NPorts * (Port1Connection - 1)) + toComplex(0.0, 1.0) * myWaveRepresentation.Vector(5 + 4 * (Port2Connection - 1) + 4 * NPorts * (Port1Connection - 1))
        '        For i1 As Integer = 1 To MechanismList1.FrequencyList.NRows
        '            S = b(i1) / a(i1)
        '            mySParamRepresentation(i1, 4) = S.Re : mySParamRepresentation(i1, 5) = S.Im
        '        Next i1
        '        'SR22 = b2m/a2m, drive on Port2Connection
        '        a = myWaveRepresentation.Vector(2 + 4 * (Port2Connection - 1) + 4 * NPorts * (Port2Connection - 1)) + toComplex(0.0, 1.0) * myWaveRepresentation.Vector(3 + 4 * (Port2Connection - 1) + 4 * NPorts * (Port2Connection - 1))
        '        b = myWaveRepresentation.Vector(4 + 4 * (Port2Connection - 1) + 4 * NPorts * (Port2Connection - 1)) + toComplex(0.0, 1.0) * myWaveRepresentation.Vector(5 + 4 * (Port2Connection - 1) + 4 * NPorts * (Port2Connection - 1))
        '        For i1 As Integer = 1 To MechanismList1.FrequencyList.NRows
        '            S = b(i1) / a(i1)
        '            mySParamRepresentation(i1, 8) = S.Re : mySParamRepresentation(i1, 9) = S.Im
        '        Next i1
        '        'SR12 = b1m/a2m, drive on Port2Connection
        '        b = myWaveRepresentation.Vector(4 + 4 * (Port1Connection - 1) + 4 * NPorts * (Port2Connection - 1)) + toComplex(0.0, 1.0) * myWaveRepresentation.Vector(5 + 4 * (Port1Connection - 1) + 4 * NPorts * (Port2Connection - 1))
        '        For i1 As Integer = 1 To MechanismList1.FrequencyList.NRows
        '            S = b(i1) / a(i1)
        '            mySParamRepresentation(i1, 6) = S.Re : mySParamRepresentation(i1, 7) = S.Im
        '        Next i1
        '    End If

        'Else

        '    'This is a switch-term measurement. Stuff the switch terms.
        '    If Port2Connection > 0 Then 'If this is also not a one-port
        '        'GR = a1m/b1m, drive on Port2Connection
        '        a = myWaveRepresentation.Vector(2 + 4 * (Port1Connection - 1) + 4 * NPorts * (Port2Connection - 1)) + toComplex(0.0, 1.0) * myWaveRepresentation.Vector(3 + 4 * (Port1Connection - 1) + 4 * NPorts * (Port2Connection - 1))
        '        b = myWaveRepresentation.Vector(4 + 4 * (Port1Connection - 1) + 4 * NPorts * (Port2Connection - 1)) + toComplex(0.0, 1.0) * myWaveRepresentation.Vector(5 + 4 * (Port1Connection - 1) + 4 * NPorts * (Port2Connection - 1))
        '        For i1 As Integer = 1 To MechanismList1.FrequencyList.NRows
        '            S = a(i1) / b(i1)
        '            mySParamRepresentation(i1, 2) = S.Re : mySParamRepresentation(i1, 3) = S.Im
        '        Next i1
        '        'GF = a2m/b2m, drive on Port1Connection
        '        a = myWaveRepresentation.Vector(2 + 4 * (Port2Connection - 1) + 4 * NPorts * (Port1Connection - 1)) + toComplex(0.0, 1.0) * myWaveRepresentation.Vector(3 + 4 * (Port2Connection - 1) + 4 * NPorts * (Port1Connection - 1))
        '        b = myWaveRepresentation.Vector(4 + 4 * (Port2Connection - 1) + 4 * NPorts * (Port1Connection - 1)) + toComplex(0.0, 1.0) * myWaveRepresentation.Vector(5 + 4 * (Port2Connection - 1) + 4 * NPorts * (Port1Connection - 1))
        '        For i1 As Integer = 1 To MechanismList1.FrequencyList.NRows
        '            S = a(i1) / b(i1)
        '            mySParamRepresentation(i1, 4) = S.Re : mySParamRepresentation(i1, 5) = S.Im
        '        Next i1
        '    End If

        'End If



        Return mySParamRepresentation

    End Function

    ''' <summary>
    ''' Form the equivalent definition of the standard
    ''' </summary>
    ''' <param name="Model">This is the model, cascade, etc. for the definition</param>
    ''' <param name="MechanismList1">The Mechanism List</param>
    ''' <returns>The scattering parameters of the equivalent definition</returns>
    ''' <remarks>The equivalent definition is determined by cascading scattering parameters.
    ''' SParEquivDef = TDUT1^-1 * TCable1 * TestPort1 * Definition * TestPort2 * TCable2 * TDUT2^-1
    ''' When the test ports are consistent, TestPort1 and TestPort2 are consistent throughout the calibration.
    ''' When flush thrus and flat shorts are used, TestPort1 and TestPort2 change, and should be worked explicitely into the standard definitions.</remarks>
    Protected Friend Function EquivalentDefinition(ByVal Model As Object, ByVal MechanismList1 As MechanismList) As RealMatrix

        'Form the scattering parameters of the model
        Dim SParEquivDef As New RealMatrix(MechanismList1.FrequencyList.NRows, 9, ".s2p") : SParEquivDef.InitializeAsSParams() : SParEquivDef.Vector(1) = MechanismList1.FrequencyList   'This will be the equivalent definition when we are done.
        Dim SPar As New RealMatrix(MechanismList1.FrequencyList.NRows, 9, ".s2p") : SPar.InitializeAsSParams() : SPar.Vector(1) = MechanismList1.FrequencyList             'Temporary variable

        'Cascade on the DUT, Cables, and Test Ports to obtain the equivalent definition
        'SParEquivDef = TDUT1^-1 * TCable1 * TestPort1 * Definition * TestPort2 * TCable2 * TDUT2^-1

        If myCurrentPort2Connection > 0 Then

            SParEquivDef = myPortModels(myCurrentPort2Connection).getSParams(MechanismList1)       'DUT Port 2
            Call SParEquivDef.ReversePorts()                                'Flip ports around on port 2
            Call SParEquivDef.Invert()                                      'We invert the DUT ports: TDUT2^-1

            SPar = myDUTPortModels(myCurrentPort2Connection).getSParams(MechanismList1)       'Test Port 2
            Call SPar.ReversePorts()
            SParEquivDef = SPar.CascadeSParameters(SParEquivDef)    'Add to the Cascade: TestPort2 * TDUT2^-1

        End If

        SPar = Model.getSParams(MechanismList1)                 'Definition
        SParEquivDef = SPar.CascadeSParameters(SParEquivDef)    'Add to the Cascade: Definition * TestPort2 * TDUT2^-1

        SPar = myPortModels(myCurrentPort1Connection).getSParams(MechanismList1)       'Test Port 1
        SParEquivDef = SPar.CascadeSParameters(SParEquivDef)    'Add to the Cascade: TestPort1 * Definition * TestPort2 * TDUT2^-1

        SPar = myDUTPortModels(myCurrentPort1Connection).getSParams(MechanismList1)       'DUT Port 1
        Call SPar.Invert()                                      'We invert the DUT ports
        SParEquivDef = SPar.CascadeSParameters(SParEquivDef)    'Add to the Cascade: TDUT1^-1 * TestPort1 * Definition * TestPort2 * TDUT2^-1

        Return SParEquivDef                                     'The equivalent definition.

    End Function

End Class

Public Class MDIF_Var_Sweep

    Public name As String
    Public start As Double
    Public finish As Double

    Public Sub New(ByVal name_in As String, ByVal start_in As Double, ByVal finish_in As Double)

        name = name_in
        start = start_in
        finish = finish_in

    End Sub

End Class

'Some additional classes and functions which I put in HPList.
'If you want to add them to HPList then strip off the module and extension tags, then remove the first mdif argument and the references
'to it in the function.

Module LaurenceHPList
    <System.Runtime.CompilerServices.Extension()>
    Public Function GetBlockIndexFromVarRanges(mdif As MDIF, ByRef VariableSpecs As MDIF_Var_Sweep()) As Integer()

        'Dim block_list As Integer() = Nothing
        Dim block_list(mdif.BlockCount() - 1) As Integer
        For i As Integer = 0 To mdif.BlockCount() - 1
            block_list(i) = i
        Next

        For Each VariableSpec As Object In VariableSpecs
            block_list = GetBlockIndicesFromVarRange(mdif, VariableSpec, block_list)
        Next

        Return block_list

    End Function

    Private Function GetBlockIndicesFromVarRangeWrapper(mdif As MDIF, ByRef VariableSpec As MDIF_Var_Sweep, ByVal block_list As Integer()) As Integer()

        ' First run?
        If block_list Is Nothing Then
            ReDim block_list(mdif.BlockCount() - 1)
        End If

        Return GetBlockIndicesFromVarRange(mdif, VariableSpec, block_list)

    End Function

    Private Function GetBlockIndicesFromVarRange(mdif As MDIF, ByRef VariableSpec As MDIF_Var_Sweep, ByVal block_list As Integer()) As Integer()

        Dim max = Math.Max(VariableSpec.start, VariableSpec.finish)
        Dim min = Math.Min(VariableSpec.start, VariableSpec.finish)
        Dim block_list_pointer As Integer = 0

        For i As Integer = 0 To block_list.Length() - 1
            For j As Integer = 0 To mdif.BlockVARs(block_list(i)).HPNames.Length - 1
                ' Find the right variable
                If mdif.BlockVARs(block_list(i)).HPNames(j) = VariableSpec.name Then
                    ' Are we in range of the sweep?
                    Dim tmp = mdif.BlockVARs(block_list(i)).GetValueDouble(j)
                    If (min <= mdif.BlockVARs(block_list(i)).GetValueDouble(j)) AndAlso (mdif.BlockVARs(i).GetValueDouble(j) <= max) Then
                        ' Yes, so add index to block_list
                        block_list(block_list_pointer) = block_list(i)
                        block_list_pointer += 1
                    End If
                End If
            Next
        Next

        ' Trim block_list
        ReDim Preserve block_list(block_list_pointer - 1)

        Return block_list

    End Function
End Module