#pragma once

#define _WIN32_DCOM

#include <windows.h>
#include <comdef.h>
#include <iostream>

#include <taskschd.h>
#pragma comment(lib, "taskschd.lib")
#pragma comment(lib, "comsupp.lib")

class Scheduler {
private:
	std::string taskName, author;
	std::string startDate, expireDate;

	HRESULT hr;
	ITaskService* pService;
	ITaskFolder* pRootFolder;
	ITaskDefinition* pTask;
	IRegistrationInfo* pRegInfo;
	ITaskSettings* pSettings;
	ITriggerCollection* pTriggerCollection;
	ITrigger* pTrigger;
	ILogonTrigger* pLogonTrigger;
	IActionCollection* pActionCollection;
	IAction* pAction;
	IExecAction* pExecAction;
	IRegisteredTask* pRegisteredTask;

	void errorExit(bool CoUninit, ITaskService* pService, ITaskFolder* pRootFolder, ITaskDefinition* pTask) {
		if (pService)
			pService->Release();

		if (pRootFolder)
			pRootFolder->Release();

		if (pTask)
			pTask->Release();

		if (CoUninit)
			CoUninitialize();

		system("PAUSE");
		exit(EXIT_FAILURE);
	}

	void init() {
		//  Initialize COM.
		hr = CoInitializeEx(0, COINIT_MULTITHREADED);
		if (FAILED(hr)) {
			printf("\nCoInitializeEx failed: %x", hr);
			errorExit(0, 0, 0, 0);
		}
	}

	void setSecurity() {
		//  Set general COM security levels.
		hr = CoInitializeSecurity(0, -1, 0, 0, RPC_C_AUTHN_LEVEL_PKT_PRIVACY, RPC_C_IMP_LEVEL_IMPERSONATE, 0, 0, 0);

		if (FAILED(hr)) {
			printf("\nCoInitializeSecurity failed: %x", hr);
			errorExit(1, 0, 0, 0);
		}
	}

	void createInstance() {
		//  Create an instance of the Task Service. 

		hr = CoCreateInstance(CLSID_TaskScheduler, 0, CLSCTX_INPROC_SERVER, IID_ITaskService, (void**)&pService);
		if (FAILED(hr)) {
			printf("Failed to create an instance of ITaskService: %x", hr);
			errorExit(1, 0, 0, 0);
		}
	}

	void connectTask() {
		//  Connect to the task service.
		hr = pService->Connect(_variant_t(), _variant_t(),
			_variant_t(), _variant_t());
		if (FAILED(hr)) {
			printf("ITaskService::Connect failed: %x", hr);
			errorExit(1, pService, 0, 0);
		}
	}

	void getRootTaskFolder() {
		//  Get the pointer to the root task folder.  This folder will hold the
		//  new task that is registered.

		hr = pService->GetFolder(_bstr_t(L"\\"), &pRootFolder);
		if (FAILED(hr)) {
			printf("Cannot get Root Folder pointer: %x", hr);
			errorExit(1, pService, 0, 0);
		}
	}

	void deleteOldTask() {
		//  If the same task exists, remove it.
		pRootFolder->DeleteTask(_bstr_t(taskName.c_str()), 0);
	}

	void setNewTask() {
		hr = pService->NewTask(0, &pTask);

		if (FAILED(hr)) {
			printf("Failed to create a task definition: %x", hr);
			errorExit(1, 0, pRootFolder, 0);
		}
	}

	void setInfo() {
		hr = pTask->get_RegistrationInfo(&pRegInfo);
		if (FAILED(hr)) {
			printf("\nCannot get identification pointer: %x", hr);
			errorExit(1, 0, pRootFolder, pTask);
		}

		hr = pRegInfo->put_Author(_bstr_t(author.c_str()));

		if (FAILED(hr)) {
			printf("\nCannot put identification info: %x", hr);
			errorExit(1, 0, pRootFolder, pTask);
		}

		hr = pTask->get_Settings(&pSettings);
		if (FAILED(hr)) {
			printf("\nCannot get settings pointer: %x", hr);
			errorExit(1, 0, pRootFolder, pTask);
		}

		//  Set setting values for the task. 
		hr = pSettings->put_DisallowStartIfOnBatteries(VARIANT_FALSE);
		hr = pSettings->put_StopIfGoingOnBatteries(VARIANT_FALSE);
		hr = pSettings->put_ExecutionTimeLimit(_bstr_t(L"PT0S"));
		hr = pSettings->put_AllowHardTerminate(VARIANT_FALSE);
		hr = pSettings->put_Hidden(VARIANT_TRUE);
		if (FAILED(hr)) {
			printf("\nCannot put setting info: %x", hr);
			errorExit(1, 0, pRootFolder, pTask);
		}

		hr = pTask->get_Triggers(&pTriggerCollection);
		if (FAILED(hr)) {
			printf("\nCannot get trigger collection: %x", hr);
			errorExit(1, 0, pRootFolder, pTask);
		}

		hr = pTriggerCollection->Create(TASK_TRIGGER_LOGON, &pTrigger);
		if (FAILED(hr)) {
			printf("\nCannot create the trigger: %x", hr);
			errorExit(1, 0, pRootFolder, pTask);
		}

		hr = pTrigger->QueryInterface(
			IID_ILogonTrigger, (void**)&pLogonTrigger);
		if (FAILED(hr)) {
			printf("\nQueryInterface call failed for ILogonTrigger: %x", hr);
			errorExit(1, 0, pRootFolder, pTask);
		}

		hr = pLogonTrigger->put_Id(_bstr_t(L"Trigger1"));
		if (FAILED(hr))
			printf("\nCannot put the trigger ID: %x", hr);
	}

	void setStartDate() {
		hr = pLogonTrigger->put_StartBoundary(_bstr_t(startDate.c_str()));
		if (FAILED(hr))
			printf("\nCannot put the start boundary: %x", hr);
	}

	void setExpireDate() {
		hr = pLogonTrigger->put_EndBoundary(_bstr_t(expireDate.c_str()));
		if (FAILED(hr))
			printf("\nCannot put the end boundary: %x", hr);
	}

	void setUser() {
		hr = pLogonTrigger->put_UserId(_bstr_t(L"DOMAIN\\UserName"));
		if (FAILED(hr)) {
			printf("\nCannot add user ID to logon trigger: %x", hr);
			errorExit(1, 0, pRootFolder, pTask);
		}
	}

	void setQI() {
		//  QI for the executable task pointer.
		hr = pAction->QueryInterface(
			IID_IExecAction, (void**)&pExecAction);

		if (FAILED(hr)) {
			printf("\nQueryInterface call failed for IExecAction: %x", hr);
			errorExit(1, 0, pRootFolder, pTask);
		}
	}

	void setPath(std::wstring wstrExecutablePath) {
		//  Set the path of the executable to notepad.exe.
		hr = pExecAction->put_Path(_bstr_t(wstrExecutablePath.c_str()));

		if (FAILED(hr)) {
			printf("\nCannot set path of executable: %x", hr);
			errorExit(1, 0, pRootFolder, pTask);
		}
	}

	void saveTask() {
		hr = pRootFolder->RegisterTaskDefinition(_bstr_t(taskName.c_str()), pTask, TASK_CREATE_OR_UPDATE, _variant_t(L"S-1-5-32-544"),
			_variant_t(), TASK_LOGON_GROUP, _variant_t(L""), &pRegisteredTask);

		if (FAILED(hr)) {
			printf("\nError saving the Task : %x", hr);
			errorExit(1, 0, pRootFolder, pTask);
		}
	}

public:

	Scheduler(std::string taskName, std::string author) {
		this->taskName = taskName;
		this->author = author;
	}

	//  format should be YYYY-MM-DDTHH:MM:SS(+-)(timezone).
	void setStartDate(std::string startDate) {
		this->startDate = startDate;
	}

	void setExpireDate(std::string expireDate) {
		this->expireDate = expireDate;
	}

	void schedule() {

		init();
		setSecurity();

		//  Get the windows directory and set the path to notepad.exe.
		std::wstring wstrExecutablePath = _wgetenv(L"WINDIR");
		wstrExecutablePath += L"\\SYSTEM32\\NOTEPAD.EXE";

		createInstance();

		connectTask();

		getRootTaskFolder();

		deleteOldTask();

		setNewTask();

		pService->Release();  // COM clean up.  Pointer is no longer used.

		setInfo();

		pRegInfo->Release();
		pSettings->Release();
		pTriggerCollection->Release();
		pTrigger->Release();

		if (!startDate.empty())
			setStartDate();

		if (!expireDate.empty())
			setExpireDate();

		pLogonTrigger->Release();

		//  Get the task action collection pointer.
		hr = pTask->get_Actions(&pActionCollection);
		if (FAILED(hr)) {
			printf("\nCannot get Task collection pointer: %x", hr);
			errorExit(1, 0, pRootFolder, pTask);
		}

		hr = pActionCollection->Create(TASK_ACTION_EXEC, &pAction);
		if (FAILED(hr)) {
			printf("\nCannot create the action: %x", hr);
			errorExit(1, 0, pRootFolder, pTask);
		}
		pActionCollection->Release();

		setQI();

		pAction->Release();

		setPath(wstrExecutablePath);

		pExecAction->Release();

		saveTask();

		printf("\n Success! Task successfully registered. ");

		// Clean up
		pRootFolder->Release();
		pTask->Release();
		pRegisteredTask->Release();
		CoUninitialize();
	}
};