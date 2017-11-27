Attribute VB_Name = "MStartup"
Option Explicit

' Copyright (c) 2017 Jason Peter Brown <jason@bitspaces.com>
'
' MIT License
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

' Default host to listen for FCGI requests
Private Const mc_DefaultHostName As String = "localhost"

' Command line parameters
Private Const mc_ParamShutdown As String = "/stop" ' Stop all running listener processes
Private Const mc_ParamSpawnCount As String = "/spawn"  ' Number of listener processes to spawn
Private Const mc_ParamListenHost As String = "/host"  ' Host name for spawned listener processes
Private Const mc_ParamListenPort As String = "/port"  ' In LISTENER mode, the port number for the current listener process
                                                      ' In SPAWNER mode, the port number to start spawning listener processes on (+1 for each subsequent process)

' FCGI listener object (one per listener process)
Private mo_FcgiServer As vbFcgiLib.CFcgiServer

' Example usage commandlines:
' Stop all running vbFcgi listener processes:

' vbfcgihost.exe /stop

' Start 4 FCGI listener processes on the localhost, starting at port 9000:
' Four processes will run, listening on ports 9000, 9001, 9002, 9003

' vbfcgihost /spawn 4 /port 9000

Sub Main()
   Dim la_Cmd() As String
   Dim l_Cmd As String
   Dim l_CmdLen As Long
   Dim ii As Long
   Dim l_Param As String
   Dim l_Value As String

   Dim l_SpawnCount As Long
   Dim l_Host As String
   Dim l_Port As Long

   Dim l_LastEvents As Double
   Dim l_LastListenerCheck As Double
   Dim l_LastShutdownCheck As Double

   Dim l_SpawnMode As Boolean
   
   Dim la_Failures() As Long
   
   On Error GoTo ErrorHandler
 
   l_Cmd = Trim$(LCase$(apiGetCommandLine))

   If l_Cmd = "" Then
      ' Command line required, raise error
      apiOutputDebugString "ASDASD"
      
      Err.Raise 5, , "Command line required."
      
   Else
      apiOutputDebugString "Command-Line: " & l_Cmd

      ' Parse command line
      Do
         ' Flatten extra whitespace
         l_CmdLen = Len(l_Cmd)

         l_Cmd = Replace$(l_Cmd, "  ", " ")
      Loop While Len(l_Cmd) <> l_CmdLen

      ' Split command line into array of key/value pairs
      la_Cmd = Split(Command, " ")

      apiOutputDebugString "Passed command elements: " & UBound(la_Cmd) + 1

      For ii = 0 To UBound(la_Cmd) Step 2
         l_Param = la_Cmd(ii)   ' Get the next parameter

         apiOutputDebugString "Found command parameter: " & l_Param

         Select Case l_Param
         Case mc_ParamShutdown
            ' Shutdown command detected, ignore all other commands and run the process shutdown process

            If ShutdownRunningListeners Then
               apiOutputDebugString "Running listeners shutdown."
            Else
               apiOutputDebugString "There was a problem trying to shutdown running listeners."
            End If

         Case Else
            ' Non-shutdown parameter found
            
            ' There are 2 acceptable command modes from here:

            ' 1) Spawner mode  - This mode spawns new listener processes.
            '                    This mode requires 2 or 3 params:
            '                    mc_ParamSpawnCount, mc_ParamListenPort, and the optional mc_ParamListenHost

            ' 2) Listener mode - This mode starts the CFCGI class and begins listening for FCGI data
            '                    This mode requires 1 or 2 params:
            '                    mc_ParamListenPort, and the optional mc_ParamListenHost
            '                    In general, this mode will not be used by the user, but is used by
            '                    Spawner mode to start the listener processes.

            l_Value = la_Cmd(ii + 1)   ' Get the value associated with the current key

            apiOutputDebugString "Found value for " & l_Param & ": " & l_Value

            Select Case l_Param
            Case mc_ParamSpawnCount
               ' Number of FCGI listener processes to spawn
               l_SpawnCount = l_Value
               l_SpawnMode = True

            Case mc_ParamListenPort
               ' The port for this listener process to use
               l_Port = l_Value

            Case mc_ParamListenHost
               ' The host name/IP address for this (or spawned) listener processes to use
               l_Host = l_Value

            Case Else
               Err.Raise 5, , "Unknown parameter: " & l_Param

            End Select
         End Select
      Next ii
   End If

   If Not MutexExists(ShutdownMutexName) Then
      If l_Host = "" Then
         ' Host is missing, use default host name
         l_Host = mc_DefaultHostName
      End If

      If l_SpawnMode Then
         ' Entering SPAWNER mode
         ' In this mode we will spawn and monitor FCGI listener processes
         
         If (l_SpawnCount > 0) And (l_Port > 0) And (l_Port < 49152) Then
            ' Spawn additional FCGI listener processes by calling path of this EXE
            ' with next port number parameter
            
            ReDim la_Failures(0 To l_SpawnCount - 1)  ' For tracking failure counts of a listener process

            ' Check if other FCGI listener processes are alrady running, and if so stop them
            If ShutdownRunningListeners Then
               ' Spawn listeners
               For ii = 0 To l_SpawnCount - 1
                  libRc5Factory.C.FSO.ShellExecute apiExePath, , , mc_ParamListenHost & " " & l_Host & " " & mc_ParamListenPort & " " & (l_Port + ii)
               Next ii
            Else
               Err.Raise vbObjectError, , "Could not shutdown existing FCGI listener processes."
            End If
         Else
            Err.Raise 5, , "Need a spawn count & port for spawner mode."
         End If
         
         ' Enter monitor loop
         ' This loop keeps the process alive and periodically checks for the existence of the listener processes
         ' If any listener processes have failed, they will be restarted
         l_LastListenerCheck = libRc5Factory.C.HPTimer
         Do
            If libRc5Factory.C.HPTimer - l_LastShutdownCheck > 0.1 Then
               If MutexExists(ShutdownMutexName) Then
                  ' Detected the shutdown mutex - exit the loop and the process will terminate
                  apiOutputDebugString "Detected shutdown mutex - leaving loop."

                  Exit Do
               End If
               l_LastShutdownCheck = libRc5Factory.C.HPTimer
            End If

            If libRc5Factory.C.HPTimer - l_LastListenerCheck > 2 Then
               If Not MutexExists(ShutdownMutexName) Then
                  ' Check if listener(s) still running
                  For ii = 0 To l_SpawnCount - 1
                     If Not MutexExists(ListenerMutexName(l_Host, l_Port + ii)) Then
                        ' Listener process must have terminated
                        ' Increment the failure count, and if over a threshold, attempt to restart the process
                        apiOutputDebugString "Mutex not found for listener on " & l_Host & ":" & (l_Port + ii) & "."
                        
                        la_Failures(ii) = la_Failures(ii) + 1
                        
                        If la_Failures(ii) > 2 Then
                           If Not MutexExists(ShutdownMutexName) Then
                              la_Failures(ii) = 0
                              
                              libRc5Factory.C.FSO.ShellExecute apiExePath, , , mc_ParamListenHost & " " & l_Host & " " & mc_ParamListenPort & " " & (l_Port + ii)
                           End If
                        End If
                     Else
                        la_Failures(ii) = 0
                     End If
                  Next ii
                  
                  l_LastListenerCheck = libRc5Factory.C.HPTimer
               End If
               
            ElseIf libRc5Factory.C.HPTimer - l_LastEvents > 0.1 Then
               DoEvents
               l_LastEvents = libRc5Factory.C.HPTimer
                                       
            Else
               apiSleep 1
            End If
            
         Loop

      Else
         ' ENTERING LISTENER MODE
         ' In this mode we will instantiate a CFcgi class and listen for FCGI data on the passed host/port.
         
         If (l_Port > 0) And (l_Port < 49152) Then
            ' Start the FCGI listener in the current process

            If CreateRunningMutex Then
               ' For shutdown detection
               apiOutputDebugString "Created ""running"" mutex."
            Else
               apiOutputDebugString "Could not create ""running"" mutex - future shutdown commands will fail."
            End If

            ' Create the FCGI listener and start listening on the appropriate host & port
            apiOutputDebugString "Creating FCGI listener on " & l_Host & ":" & l_Port

            Set mo_FcgiServer = libRc5Factory.RegFree.GetInstanceEx(pathBin & "vbFcgiLib.dll", "CFcgiServer")
            mo_FcgiServer.StartListening l_Host, l_Port

            If Not CreateListenerMutex(l_Host, l_Port) Then
               Err.Raise vbObjectError, , "Could not create listener mutex for " & l_Host & ":" & l_Port
            End If

            apiOutputDebugString "Created FCGI listener on " & l_Host & ":" & l_Port

            ' Enter the main loop
            Do
               If libRc5Factory.C.HPTimer - l_LastShutdownCheck > 0.5 Then
                  If MutexExists(ShutdownMutexName) Then
                     ' Detected the shutdown mutex - exit the loop and the process will terminate
                     apiOutputDebugString "Detected shutdown mutex - leaving loop."

                     Exit Do
                  End If
                  l_LastShutdownCheck = libRc5Factory.C.HPTimer
               End If

               If libRc5Factory.C.HPTimer - l_LastEvents > 0.1 Then
                  DoEvents
                  l_LastEvents = libRc5Factory.C.HPTimer
               Else
                  apiSleep 1
               End If
            Loop

         Else
            Err.Raise 5, , "Invalid port for listener mode: " & l_Port
         End If
      End If
   End If

Cleanup:
   On Error Resume Next
   
   If Not mo_FcgiServer Is Nothing Then
      mo_FcgiServer.StopListening
      Set mo_FcgiServer = Nothing
   End If
   
   libRc5Factory.C.CleanupRichClientDll

   Exit Sub

ErrorHandler:
   apiOutputDebugString "Error: " & Err.Number & " " & Err.Description

   Resume Cleanup
End Sub

Private Function ShutdownRunningListeners() As Boolean
   ' Attempt to shutdown all running listeners by creat the Shutdown mutex to signal them
   
   Dim l_LoopTime As Double
   Dim l_StartedLoopAt As Double

   If MutexExists(RunningMutexName) Then
      ' At least one spawned instance is running
      
      If CreateShutdownMutex Then
         ' Wait for spawned processes to shutdown, then leave
         apiOutputDebugString "Created shutdown mutex."
         
         l_StartedLoopAt = libRc5Factory.C.HPTimer
         Do
            apiOutputDebugString "Waiting for processes to shutdown."
            
            DoEvents
            apiSleep 250
            
            l_LoopTime = libRc5Factory.C.HPTimer - l_StartedLoopAt
         Loop While MutexExists(RunningMutexName) And (l_LoopTime < 60)
      
         If Not MutexExists(RunningMutexName) Then
            ShutdownRunningListeners = True
         End If
      Else
         apiOutputDebugString "Could not create shutdown mutex."
      End If
   Else
      ShutdownRunningListeners = True
   End If
End Function

Private Function CreateShutdownMutex() As Boolean
   Dim l_Mutex As Long
   
   ' Creating this mutex will signal all other FCGI listener processes to shutdown
   
   l_Mutex = apiCreateMutex(ByVal 0&, 1&, StrPtr(ShutdownMutexName))
   CreateShutdownMutex = (l_Mutex <> 0)
End Function

Private Function CreateRunningMutex() As Boolean
   Dim l_Mutex As Long
   
   ' Create this mutex before entering the wait loop to allow the existence of FCGI listeners
   ' to be detected by the shutdown mechanism
   
   l_Mutex = apiCreateMutex(ByVal 0&, 0&, StrPtr(RunningMutexName))
   CreateRunningMutex = (l_Mutex <> 0)
End Function

Private Function CreateListenerMutex(ByVal p_Host As String, ByVal p_Port As Long) As Boolean
   Dim l_Mutex As Long
   
   ' Create this mutex after the listener has been started
   ' It will be tested for when in SPAWNER mode to make sure the listener hasn't crashed
   
   l_Mutex = apiCreateMutex(ByVal 0&, 0&, StrPtr(ListenerMutexName(p_Host, p_Port)))
   CreateListenerMutex = (l_Mutex <> 0)
End Function

Private Function MutexExists(ByVal p_MutexName As String) As Boolean
   Dim l_Mutex As Long
   Dim l_LastDllError As Long
   
   ' Return true if the passed mutex name exists, otherwise false

   l_Mutex = apiOpenMutex(SYNCHRONIZE, 0, StrPtr(p_MutexName))
   l_LastDllError = Err.LastDllError
   
   If l_Mutex <> 0 Then
      apiCloseHandle l_Mutex: l_Mutex = 0
      
      MutexExists = True
   Else
      If l_LastDllError <> ERROR_FILE_NOT_FOUND Then
         Err.Raise vbObjectError, , "Unexpected error. Last DLL Error #" & l_LastDllError
      End If
   End If
End Function

Private Function ShutdownMutexName()
   Static s_Name As String
   
   If Len(s_Name) = 0 Then
      s_Name = "vbFcgi_Shutdown_" & libCrypt.SHA256(LCase$(apiExePath), True)
   End If
   
   ShutdownMutexName = s_Name
End Function

Private Function RunningMutexName()
   Static s_Name As String
   
   If Len(s_Name) = 0 Then
      s_Name = "vbFcgi_Running_" & libCrypt.SHA256(LCase$(apiExePath), True)
   End If
   
   RunningMutexName = s_Name
End Function

Private Function ListenerMutexName(ByVal p_Host As String, ByVal p_Port As Long)
   Static so_Names As vbRichClient5.cCollection
   
   Dim l_Key As String
   
   If so_Names Is Nothing Then Set so_Names = libRc5Factory.C.Collection
   
   l_Key = p_Host & ":" & p_Port
   If so_Names.Exists(l_Key) Then
      ListenerMutexName = so_Names.Item(l_Key)
   Else
      ListenerMutexName = "vbFcgi_Listen_" & libCrypt.SHA256(LCase$(apiExePath & "|" & p_Host) & "|" & p_Port, True)
      
      so_Names.Add CStr(ListenerMutexName), l_Key
   End If
End Function
