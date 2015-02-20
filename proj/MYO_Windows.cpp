////////////////////////////////////////////////////////////////////////////////////////////
// Special thanks to Mohit Arora for the tutorial on creating a simple windows service.
//
// All MYO Libraries Copyright (C) 2013-2014 Thalmic Labs Inc.
////////////////////////////////////////////////////////////////////////////////////////////


#define _USE_MATH_DEFINES
#include <cmath>
#include <iostream>
#include <windows.h>
#include <tchar.h>
#include <strsafe.h>
#include <iomanip>
#include <stdexcept>
#include <string>
#include <algorithm>
#include <myo/myo.hpp>
#pragma comment(lib, "advapi32.lib")

SERVICE_STATUS        g_ServiceStatus = {0}; 
SERVICE_STATUS_HANDLE g_StatusHandle = NULL; 
HANDLE                g_ServiceStopEvent = INVALID_HANDLE_VALUE; 
VOID WINAPI ServiceMain (DWORD argc, LPTSTR *argv);
VOID WINAPI ServiceCtrlHandler (DWORD);
DWORD WINAPI ServiceWorkerThread (LPVOID lpParam); 
#define SERVICE_NAME  _T("MYOWindows")



///////////////////////////////////////////////////////////////////////////
// MYO Specific Classes, Probably Need To Port To Library
///////////////////////////////////////////////////////////////////////////
class DataCollector : public myo::DeviceListener {

public:
    DataCollector()
    : onArm(false), isUnlocked(false), roll_w(0), pitch_w(0), yaw_w(0), currentPose()
    {
    }

    void onUnpair(myo::Myo* myo, uint64_t timestamp)
    {
		// Clean up
        roll_w = 0;
        pitch_w = 0;
        yaw_w = 0;
        onArm = false;
        isUnlocked = false;
    }

    void onOrientationData(myo::Myo* myo, uint64_t timestamp, const myo::Quaternion<float>& quat)
    {
        using std::atan2;
        using std::asin;
        using std::sqrt;
        using std::max;
        using std::min;

        float roll = atan2(2.0f * (quat.w() * quat.x() + quat.y() * quat.z()),
                           1.0f - 2.0f * (quat.x() * quat.x() + quat.y() * quat.y()));
        float pitch = asin(max(-1.0f, min(1.0f, 2.0f * (quat.w() * quat.y() - quat.z() * quat.x()))));
        float yaw = atan2(2.0f * (quat.w() * quat.z() + quat.x() * quat.y()),
                        1.0f - 2.0f * (quat.y() * quat.y() + quat.z() * quat.z()));

        roll_w = static_cast<int>((roll + (float)M_PI)/(M_PI * 2.0f) * 18);
        pitch_w = static_cast<int>((pitch + (float)M_PI/2.0f)/M_PI * 18);
        yaw_w = static_cast<int>((yaw + (float)M_PI)/(M_PI * 2.0f) * 18);
    }

    void onPose(myo::Myo* myo, uint64_t timestamp, myo::Pose pose)
    {
        currentPose = pose;

        if (pose != myo::Pose::unknown && pose != myo::Pose::rest) {
            myo->unlock(myo::Myo::unlockHold);
			myo->notifyUserAction();
        } else {
            myo->unlock(myo::Myo::unlockTimed); // Inactive
        }
    }

    void onArmSync(myo::Myo* myo, uint64_t timestamp, myo::Arm arm, myo::XDirection xDirection)
    {
        onArm = true;
        whichArm = arm;
    }

    void onArmUnsync(myo::Myo* myo, uint64_t timestamp)
    {
        onArm = false;
    }

    void onUnlock(myo::Myo* myo, uint64_t timestamp)
    {
        isUnlocked = true;
    }

    void onLock(myo::Myo* myo, uint64_t timestamp)
    {
        isUnlocked = false;
    }
	
	void print()
    {
        std::cout << '\r';

        std::cout << '[' << std::string(roll_w, '*') << std::string(18 - roll_w, ' ') << ']'
                  << '[' << std::string(pitch_w, '*') << std::string(18 - pitch_w, ' ') << ']'
                  << '[' << std::string(yaw_w, '*') << std::string(18 - yaw_w, ' ') << ']';

        if (onArm) {
            std::string poseString = currentPose.toString();

            std::cout << '[' << (isUnlocked ? "unlocked" : "locked  ") << ']'
                      << '[' << (whichArm == myo::armLeft ? "L" : "R") << ']'
                      << '[' << poseString << std::string(14 - poseString.size(), ' ') << ']';
        } else {
            std::cout << '[' << std::string(8, ' ') << ']' << "[?]" << '[' << std::string(14, ' ') << ']';
        }

        std::cout << std::flush;
    }

    bool onArm;
    myo::Arm whichArm;
    bool isUnlocked;
    int roll_w, pitch_w, yaw_w;
    myo::Pose currentPose;
	std::string poseString;
};


///////////////////////////////////////////////////////////////////////////
// Application Entry
///////////////////////////////////////////////////////////////////////////
int _tmain (int argc, TCHAR *argv[])
{
    SERVICE_TABLE_ENTRY ServiceTable[] = 
    {
        {SERVICE_NAME, (LPSERVICE_MAIN_FUNCTION) ServiceMain},
        {NULL, NULL}
    };
 
    if (StartServiceCtrlDispatcher (ServiceTable) == FALSE)
    {
        return GetLastError ();
    }
 
    return 0;
}


///////////////////////////////////////////////////////////////////////////
// Service Entry
///////////////////////////////////////////////////////////////////////////
VOID WINAPI ServiceMain (DWORD argc, LPTSTR *argv)
{
    DWORD Status = E_FAIL;
 
    // Register handler with the SCM
    g_StatusHandle = RegisterServiceCtrlHandler (SERVICE_NAME, ServiceCtrlHandler);
 
    if (g_StatusHandle == NULL) 
    {
        goto EXIT;
    }
 
    // Communicate Stop
    ZeroMemory (&g_ServiceStatus, sizeof (g_ServiceStatus));
    g_ServiceStatus.dwServiceType = SERVICE_WIN32_OWN_PROCESS;
    g_ServiceStatus.dwControlsAccepted = 0;
    g_ServiceStatus.dwCurrentState = SERVICE_START_PENDING;
    g_ServiceStatus.dwWin32ExitCode = 0;
    g_ServiceStatus.dwServiceSpecificExitCode = 0;
    g_ServiceStatus.dwCheckPoint = 0;
 
    if (SetServiceStatus (g_StatusHandle , &g_ServiceStatus) == FALSE)
    {
        OutputDebugString(_T(
          "My Sample Service: ServiceMain: SetServiceStatus returned error"));
    }
 
    // Implement ServicesStartEvent 
 
    g_ServiceStopEvent = CreateEvent (NULL, TRUE, FALSE, NULL);
    if (g_ServiceStopEvent == NULL) 
    {   
        g_ServiceStatus.dwControlsAccepted = 0;
        g_ServiceStatus.dwCurrentState = SERVICE_STOPPED;
        g_ServiceStatus.dwWin32ExitCode = GetLastError();
        g_ServiceStatus.dwCheckPoint = 1;
 
        if (SetServiceStatus (g_StatusHandle, &g_ServiceStatus) == FALSE)
	{
	    OutputDebugString(_T(
	      "MYOWindows: ServiceMain: SetServiceStatus returned error"));
	}
        goto EXIT; 
    }    
    
    // Communicate Start
    g_ServiceStatus.dwControlsAccepted = SERVICE_ACCEPT_STOP;
    g_ServiceStatus.dwCurrentState = SERVICE_RUNNING;
    g_ServiceStatus.dwWin32ExitCode = 0;
    g_ServiceStatus.dwCheckPoint = 0;
 
    if (SetServiceStatus (g_StatusHandle, &g_ServiceStatus) == FALSE)
    {
        OutputDebugString(_T(
          "MYOWindows: ServiceMain: SetServiceStatus returned error"));
    }
 
    // Start worker thread
    HANDLE hThread = CreateThread (NULL, 0, ServiceWorkerThread, NULL, 0, NULL);
   
    // Wait for worker stop
    WaitForSingleObject (hThread, INFINITE);
   
    CloseHandle (g_ServiceStopEvent);
 
    // Communicate stop
    g_ServiceStatus.dwControlsAccepted = 0;
    g_ServiceStatus.dwCurrentState = SERVICE_STOPPED;
    g_ServiceStatus.dwWin32ExitCode = 0;
    g_ServiceStatus.dwCheckPoint = 3;
 
    if (SetServiceStatus (g_StatusHandle, &g_ServiceStatus) == FALSE)
    {
        OutputDebugString(_T(
          "My Sample Service: ServiceMain: SetServiceStatus returned error"));
    }
    
EXIT:
    return;
} 


///////////////////////////////////////////////////////////////////////////
// Service Control
///////////////////////////////////////////////////////////////////////////
VOID WINAPI ServiceCtrlHandler (DWORD CtrlCode)
{
    switch (CtrlCode) 
	{
     case SERVICE_CONTROL_STOP :
 
        if (g_ServiceStatus.dwCurrentState != SERVICE_RUNNING)
           break;
 
        // Implement ServicesStopEvent Criteria
        
        g_ServiceStatus.dwControlsAccepted = 0;
        g_ServiceStatus.dwCurrentState = SERVICE_STOP_PENDING;
        g_ServiceStatus.dwWin32ExitCode = 0;
        g_ServiceStatus.dwCheckPoint = 4;
 
        if (SetServiceStatus (g_StatusHandle, &g_ServiceStatus) == FALSE)
        {
            OutputDebugString(_T(
              "MYOWindows: ServiceCtrlHandler: SetServiceStatus returned error"));
        }
 
        // Communicate shutdown of worker thread
        SetEvent (g_ServiceStopEvent);
 
        break;
 
     default:
         break;
    }
}  


///////////////////////////////////////////////////////////////////////////
// Worker 
///////////////////////////////////////////////////////////////////////////
DWORD WINAPI ServiceWorkerThread (LPVOID lpParam)
{
    while (WaitForSingleObject(g_ServiceStopEvent, 0) != WAIT_OBJECT_0)
    {                
        try 
		{
			myo::Hub hub("com.example.hello-myo");

			myo::Myo* myo = hub.waitForMyo(10000);

			if (!myo) 
			{
				throw std::runtime_error("");
			}

			std::cout << "Connected." << std::endl << std::endl;

			DataCollector collector;

			hub.addListener(&collector);

			while (1) 
			{
				hub.run(1000/20);  // 20 times a second
			}
		} 
		catch (const std::exception& e) 
		{
			return 1;
		}	
    } 
    return ERROR_SUCCESS;
} 

