Imports System.IO

Module Module1
    '------------------------------------------------------------
    '-                File Name : Module1.vb                    - 
    '-                Part of Project: Assignment3              -
    '------------------------------------------------------------
    '-                Written By: Benjamin Neeb                 -
    '-                Written On: February 1, 2021              -
    '------------------------------------------------------------
    '- File Purpose:                                            -
    '-                                                          -
    '- This file contains the main Sub for the console          -
    '- application. The user will input all their data in this  - 
    '- file. The error handling, file I/O, calculations, and    -
    '- output are performed in this file. 
    '------------------------------------------------------------
    '- Program Purpose:                                         -
    '-                                                          -
    '- This program gathers all text from a text file of the    -
    '- user's choice. It then performs statistical analysis on  -
    '- the text contained in the file. After analysis is        -
    '- complete, the program writes a generated report out to   -
    '- another text file of the user's choice. The user is then -
    '- presented the option to view the generated report from   -
    '- the console window. If the user selects 'yes', the       -
    '- output is fetched from the newly written file and        =
    '- presented on the console window. 
    '------------------------------------------------------------
    '- Global Variable Dictionary (alphabetically):             -
    '- (None)                                                   –
    '------------------------------------------------------------

    '---------------------------------------------------------------------------------------
    '--- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS ---
    '--- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS ---
    '--- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS ---
    '---------------------------------------------------------------------------------------

    Const intALPHABETNUM As Integer = 25            'Integer for array sizes
    Const strHISTCHAR As String = "*"               'String character for histogram representation
    Const chrPOSITIVE As Char = Chr(89)             'Char to look for positive response
    Const intDUPNUMBER As Integer = 50              'Integer for repeat count
    Const intMOSTINITIAL As Integer = 0             'Integer to initalize most variable
    Const intLEASTINITIAL As Integer = 2147483647   'Integer to initalize least variable
    Const intASCII_A As Integer = 65                'Integer for ASCII A
    Const intLEADINGZEROS As Integer = 2            'Integer for leading zero count

    '-----------------------------------------------------------------------------------
    '--- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS ---
    '--- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS ---
    '--- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS ---
    '-----------------------------------------------------------------------------------

    Sub Main()
        '------------------------------------------------------------
        '-            Subprogram Name: Main                         -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: January 28, 2021              -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine sets the console's attributesm, gathers  -
        '- all input from the user, and terminates the program. No  -
        '- calculations, file I/O, or error handling is performed   -
        '- in this Sub.                                             -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- (None)                                                   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- strAnalysis:     String variable which holds the entire  -
        '-                  statistical analysis that was performed -
        '- strFileContents: String variable which holds the entire  -
        '-                  contents of the user's file chosen to   -
        '-                  be statistically analyzed.              -
        '- strHeader:       String variable which holds the header  -
        '-                  message telling the user that the       -
        '-                  statistical analysis is displayed below -
        '- strWritePath:    String variable which holds the path to -
        '-                  the text file that the statistical      -
        '-                  analysis will be written to and read    -
        '-                  from                                    -
        '------------------------------------------------------------

        'Set background color to dark blue
        Console.BackgroundColor = ConsoleColor.DarkBlue
        'Clear contents on console to reset as blue
        Console.Clear()
        'Set title
        Console.Title = "Character Analysis Profiler Application"
        'Set text color as yellow
        Console.ForegroundColor = ConsoleColor.Yellow

        'Ask user for path of file
        Console.WriteLine("Please enter the path and name of the file to process:")

        'Get the contents of the file and store in a string
        Dim strFileContents As String = GetFile(Console.ReadLine())

        'If the read did not fail (returned a string), continue program
        If Not (strFileContents.Equals("")) Then
            'If file read did not fail, continue to function for statistical analysis
            Dim strAnalysis As String = StatAnalysis(strFileContents)

            'Simply adding a line of space
            Console.WriteLine()
            Console.WriteLine("Processing Completed...")

            'Simply adding a line of space
            Console.WriteLine()
            Console.WriteLine("Please enter the path And name of the report file to generate")

            'Get the destination filepath from user. Putting this in a string variable for later use.
            Dim strWritePath As String = Console.ReadLine()

            'Send path and data to a function to create the file and populate it
            If WriteFile(strWritePath, strAnalysis) Then
                'If the write operation was a success, continue program

                'Simply adding a line of space
                Console.WriteLine()
                Console.WriteLine("Report File Generation Completed...")

                'Simply adding a line of space
                Console.WriteLine()
                Console.WriteLine("Would you Like to see the report file? [Y/n]")

                'Simply adding a line of space
                Console.WriteLine()

                'If the user enters either Y or y, then show them the data (compare to capital 'Y'). Anything else will be regarded as a 'no'. 
                If Console.ReadLine().ToUpper.Equals(chrPOSITIVE) Then
                    'Center message in the middle of console window
                    Dim strHeader As String = "Character Analysis Statistics"
                    Console.WriteLine(String.Format("{0, " & Math.Floor((Console.WindowWidth / 2) + (strHeader.Length / 2)) & "}", strHeader))

                    'Simply adding a line of space
                    Console.WriteLine()

                    'Write data to console here!
                    Console.Write(GetFile(strWritePath))
                End If

                'Simply adding a line of space
                Console.WriteLine()
                Console.WriteLine("Application has completed. Press any key to end.")
            End If
        End If

        'Ending program here
        Console.ReadKey()
    End Sub

    Private Function GetFile(strPath As String) As String
        '------------------------------------------------------------
        '-                Function Name: GetFile                    -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: January 28, 2021              -
        '------------------------------------------------------------
        '- Function Purpose:                                        -
        '-                                                          -
        '- This function accepts the file path that the user wants  -
        '- to read from. The function then ensures that the file    -
        '- be successfully read from without errors. If not, the    -
        '- function takes care of error handling and will return an -
        '- empty String. If the file read is successful, the        -
        '- function reads the file's contents into a String         -
        '- variable and returns the String.                         -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- strPath:         String variable that contains the path  -
        '-                  of the file the user wants to read from -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- strContents:     String variable that contains the       -
        '-                  contents of the read file.              -
        '------------------------------------------------------------
        '- Returns:                                                 -
        '- String:          Contains the file contents              -
        '------------------------------------------------------------

        'String for file contents
        Dim strContents As String = ""

        'Must get data from an existing text file as outlined in assignment constraints
        If Not (strPath.EndsWith(".txt") Or File.Exists(strPath)) Then
            Console.WriteLine()
            Console.WriteLine(StrDup(intDUPNUMBER, "*"))
            Console.WriteLine("*** Must input data from a text (.txt) file! ***")
            Console.WriteLine(StrDup(intDUPNUMBER, "*"))
            Console.WriteLine()
            Console.WriteLine("*** Application will exit -- press any key... ***")
        Else
            Try
                'Read the contents of the file into strContents
                strContents = File.ReadAllText(strPath)
            Catch ex As Exception
                'If this failed, alert user
                Console.WriteLine()
                Console.WriteLine(StrDup(intDUPNUMBER, "*"))
                Console.WriteLine("*** Could not open input file for processing! ***")
                Console.WriteLine(StrDup(intDUPNUMBER, "*"))
                Console.WriteLine()
                Console.WriteLine("*** Application will exit -- press any key... ***")
            End Try
        End If

        'Return strContents and cast to uppercase. It will either contain the contents of the file or will equal ""
        Return strContents.ToUpper
    End Function

    Private Function StatAnalysis(strData As String) As String
        '------------------------------------------------------------
        '-                Function Name: StatAnalysis               -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: January 29, 2021              -
        '------------------------------------------------------------
        '- Function Purpose:                                        -
        '-                                                          -
        '- This function accepts the data as a String that the user -
        '- wants to analyze. The function creates a variable that   -
        '- will contain the user's statistical analysis. The        -
        '- then calls other functions to consequetevely add to the  -
        '- analysis String.
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- strData:         String variable that contains the data  -
        '-                  read from the file to be statistically  -
        '-                  analyzed                                -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- strAnalysis:     String variable that contains the       -
        '-                  entire statistical analysis that will   -
        '-                  be returned                             -
        '- dblAvg:          Double variable that contains the mean  -
        '-                  letter occurance.                       -
        '- arrChars:        Char array variable that contains       -
        '-                  letters A through Z. Populated as they  -
        '-                  are found in the file's String          -
        '- arrCount:        Integer array variable that contains    -
        '-                  the occurance count of each letter from -
        '-                  the file's String                       -
        '- intLeast:        Integer variable to hold the count of   -
        '-                  the letter that occurs the least amount -
        '-                  of times                                -
        '- intMost:         Integer variable to hold the count of   -
        '-                  the letter that occurs the most amount  -
        '-                  of times                                -
        '------------------------------------------------------------
        '- Returns:                                                 -
        '- String:          Contains the entire statistical         -
        '-                  analysis                                -
        '------------------------------------------------------------

        Dim strAnalysis As String = ""

        'Integer to hold the most occurance of letters
        Dim intMost As Integer = intMOSTINITIAL

        'Integer to hold the least occurance of letters
        Dim intLeast As Integer = intLEASTINITIAL

        'Create a String array of all letters
        Dim arrChars(intALPHABETNUM) As Char
        'Create an Integer array of all letters occurances
        Dim arrCount(intALPHABETNUM) As Integer

        'Create a double variable for the average letter count occured
        Dim dblAvg As Double = MakeArray(strData, arrChars, arrCount, strAnalysis, intMost, intLeast)

        'Return the entire analysis consisting of a String graph of all letters and their occurances
        Return strAnalysis & MakeString(dblAvg, arrChars, arrCount, intMost, intLeast)
    End Function

    Private Function MakeArray(strData As String, ByRef arrChars As Char(), ByRef arrCount As Integer(), ByRef strAnalysis As String, ByRef intMost As Integer, ByRef intLeast As Integer) As Double
        '------------------------------------------------------------
        '-                Function Name: MakeArray                  -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: January 29, 2021              -
        '------------------------------------------------------------
        '- Function Purpose:                                        -
        '-                                                          -
        '- This function populates the arrays that will hold the    -
        '- primary data for statistical analysis. The arrays will   -
        '- hold the chars as they are encountered in the file's     -
        '- String and the count of the letters as they are          -
        '- encountered in the file's String, respectively. The      -
        '- function then returns the average letter occurance as a  -
        '- Double.                                                  -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- strAnalysis:     String variable passed ByRef that       -
        '-                  contains the consequetevely built       -
        '-                  statistical analysius                   -
        '- arrChars:        Char array passed ByRef that will be    -
        '-                  populated with the chars as they appear -
        '-                  in the file's String                    -
        '- arrCount:        Integer array passed ByRef that         -
        '-                  contains the occurance count of each    -
        '-                  letter from the file's String           -
        '- strData:         String variable that contains the file  -
        '-                  contents that was read for statistical  -
        '-                  analysis                                -
        '- intLeast:        Integer variable passed ByRef that will -
        '-                  hold the count of the letter that       -
        '-                  occurs the least amount of times        -
        '- intMost:         Integer variable passed ByRef that will -
        '-                  hold the count of the letter that       -
        '-                  occurs the most amount of times         -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- intASCII:        Integer variable that holds the ASCII   -
        '-                  value of the current character. Used    -
        '-                  to find the corresponding array index   -
        '-                  for incrementing the letter's occurance -
        '- intCtrLetter:    Integer variable that counts the        -
        '-                  occurance of letters in the file's      -
        '-                  String                                  -                  
        '- intCtrOther:     Integer variable that counts the        -
        '-                  occurance of non-letters in the file's  -
        '-                  String                                  -
        '- blnFound:        Boolean variable that flags whether or  -
        '-                  not the character was found in the      -
        '-                  array                                   -
        '------------------------------------------------------------
        '- Returns:                                                 -
        '- Double:          Contains the average letter occurance   -
        '------------------------------------------------------------

        'Create counters for letters and other characters
        Dim intCtrLetter As Integer = 0
        Dim intCtrOther As Integer = 0

        'Go through entire alphabet and get count of each letter
        For Each c As Char In strData
            'If the currect character is a letter...
            If Char.IsLetter(c) Then

                'Create boolean flag that will be used to indicate if we found the character in the array
                Dim blnFound As Boolean = False

                'Cycle through entire array in search for the current character
                For i As Integer = 0 To intALPHABETNUM
                    'If the array element is null or empty, skip
                    If Not (String.IsNullOrEmpty(arrChars(i))) Then
                        'If the character is found in the current array element...
                        If arrChars(i).Equals(c) Then
                            'Increment the count
                            arrCount.SetValue(arrCount(i) + 1, i)

                            'Keep track of the count of the most and least occuring letter
                            If arrCount(i) > intMost Then
                                'Assign new most
                                intMost = arrCount(i)
                            ElseIf arrCount(i) < intLeast Then
                                'Assign new least
                                intLeast = arrCount(i)
                            End If

                            'Set found flag to true
                            blnFound = True

                            'Since we found the letter, exit the loop
                            Exit For
                        End If
                    End If
                Next

                If Not (blnFound) Then
                    'If we did not find the character in the array, add it
                    'Get ASCII value of character
                    Dim intASCII As Integer = Convert.ToByte(c)

                    'We are setting the characters in alphabetic order in the array
                    'To do this, subtract 65 from the ASCII value and set the value to that index
                    arrChars.SetValue(c, intASCII - intASCII_A)
                    arrCount.SetValue(1, intASCII - intASCII_A)
                End If

                'Increment letter counter
                intCtrLetter += 1
            Else
                'Increment other character counter
                intCtrOther += 1
            End If
        Next

        'Now, add the characters that did not exist in the input file followed by a zero
        For i As Integer = 0 To intALPHABETNUM
            If Not (Char.IsLetter(arrChars(i))) Then
                'Add ASCII value (element index + 65)
                arrChars.SetValue(Chr(i + intASCII_A), i)
                arrCount.SetValue(0, i)
                'If one of these cases existed, update the lowest occurance to zero
                intLeast = 0
            End If
        Next

        'Write initial lines of data to string
        strAnalysis = "There were a total of " & intCtrLetter + intCtrOther & " characters processed." & vbCrLf
        strAnalysis = strAnalysis & intCtrLetter & " were letters and " & intCtrOther & " were other characters." & vbCrLf
        strAnalysis = strAnalysis & vbCrLf & ""

        'Return the average letter value
        Return intCtrLetter / (intALPHABETNUM + 1)
    End Function

    Private Function MakeString(dblAvg As Double, arrChars As Char(), arrCount As Integer(), intMost As Integer, intLeast As Integer)
        '------------------------------------------------------------
        '-                Function Name: MakeString                 -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: January 29, 2021              -
        '------------------------------------------------------------
        '- Function Purpose:                                        -
        '-                                                          -
        '- This function populates the remainder of the statistical -
        '- analysis String. This is comprised of the histogram      -
        '- graph and final analysis data pieces. The function then  -
        '- returns the statistical analysis in its entirety         -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- dblAvg:          Double variable contains the count of   -
        '-                  the average letter occurance            -
        '- arrChars:        Char array that will be populated with  -
        '-                  the chars as they appear in the file's  -
        '-                  String                                  -
        '- arrCount:        Integer array that contains the         -
        '-                  occurance count of each letter from the -
        '-                  file's String                           -
        '- intLeast:        Integer variable that will hold the     -
        '-                  count of the letter that occurs the     -
        '-                  least amount of times                   -
        '- intMost:         Integer variable that will hold the     -
        '-                  count of the letter that occurs the     -
        '-                  most amount of times                    -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- intFormula:      Integer variable used to create the     -
        '-                  histogram representing letters and      -
        '-                  their occurances                        -
        '- strGraph:        String variable used to create the      -
        '-                  histogram representing letters and      -
        '-                  their occurances                        -
        '- intHistSize:     Integer variable used to size the       -
        '-                  histogram to the conole's window        -
        '- strLeast:        String variable that announces which    -
        '-                  letter(s) occurred the least times      -
        '- strMost:         String variable that announces which    -
        '-                  letter(s) occurred the most times       -
        '- strPadding:      String variable that adds zeros to the  -
        '-                  front of each letter's occurance        -
        '------------------------------------------------------------
        '- Returns:                                                 -
        '- String:          Contains the file contents              -
        '------------------------------------------------------------

        'Will hold this subs contribution to the overall output
        Dim strGraph As String = ""

        'Will hold the data for which letter occured the most often
        Dim strMost As String = ""

        'Will hold the data for which letter occured the least often
        Dim strLeast As String = ""

        'Determine zero padding. Make it (num of digits) of max occurance plus leading zeros. 
        Dim strPadding As String = String.Format(StrDup(intMost.ToString.Length + intLEADINGZEROS, "0"))

        'Determine histogram length as 1/4 of the current windows size
        Dim intHistSize As Integer = Math.Floor(Console.WindowWidth / 4)

        'Loop through each letter in the char array
        For i As Integer = 0 To intALPHABETNUM
            'If this letter occured the most or least, record it
            If arrCount(i) = intMost Then
                If strMost.Equals("") Then
                    'Add first occurance of most letter
                    strMost = "Highest Letter Utilization: " & intMost & " on " & arrChars(i)
                Else
                    'Add multiple occurances of most letters
                    strMost = strMost & ", " & arrChars(i)
                End If
            ElseIf arrCount(i) = intLeast Then
                If strLeast.Equals("") Then
                    'Add first occurance of least letter
                    strLeast = "Lowest Letter Utilization: " & intLeast & " on " & arrChars(i)
                Else
                    'Add multiple occurances of least letters
                    strLeast = strLeast & ", " & arrChars(i)
                End If
            End If

            'Create letter's histogram size
            Dim intFormula As Integer = Math.Ceiling(arrCount(i) / intMost * intHistSize)

            'Create the line for the letter (Letter : Occurance Graph)
            strGraph = strGraph & arrChars(i) & " : " & arrCount(i).ToString(strPadding) & " " & String.Format(StrDup(intFormula, strHISTCHAR)) & vbCrLf
        Next

        'Add final lines to statistical analysis
        strGraph = strGraph & vbCrLf & ""
        strGraph = strGraph & "Average Letter Value: " & dblAvg & vbCrLf
        strGraph = strGraph & strMost & vbCrLf
        strGraph = strGraph & strLeast & vbCrLf

        Return strGraph
    End Function

    Private Function WriteFile(strPath As String, strData As String) As Boolean
        '------------------------------------------------------------
        '-                Function Name: WriteFile                  -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: January 28, 2021              -
        '------------------------------------------------------------
        '- Function Purpose:                                        -
        '-                                                          -
        '- This function accepts the file path and String that the  -
        '- user wants to write to. The function then ensures that   -
        '- the file can be successfully written to without errors.  -
        '- If not, the function takes care of error handling and    -
        '- will return false. If the file write is successful, the  -
        '- function writes the String to the file and returns true  -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- strData:         String variable that contains data that -
        '-                  will be written to the file             -
        '- strPath:         String variable that contains the path  -
        '-                  of the file the user wants to read from -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------
        '- Returns:                                                 -
        '- Boolean:         File write success/failure              -
        '------------------------------------------------------------

        'Must write to a text file as outlined in assignment constraints
        If Not (strPath.EndsWith(".txt")) Then
            Console.WriteLine()
            Console.WriteLine(StrDup(intDUPNUMBER, "*"))
            Console.WriteLine("*** Must output analysis to a text (.txt) file! ***")
            Console.WriteLine(StrDup(intDUPNUMBER, "*"))
            Console.WriteLine()
            Console.WriteLine("*** Application will exit -- press any key... ***")
        Else
            Try
                'Write the analysis out to a text file
                File.WriteAllText(strPath, strData)
                Return True
            Catch ex As Exception
                'If this failed, alert user
                Console.WriteLine()
                Console.WriteLine(StrDup(intDUPNUMBER, "*"))
                Console.WriteLine("*** Could not open output file path for writing! ***")
                Console.WriteLine(StrDup(intDUPNUMBER, "*"))
                Console.WriteLine()
                Console.WriteLine("*** Application will exit -- press any key... ***")
            End Try
        End If

        Return False
    End Function
End Module
