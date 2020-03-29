#include <ShlObj.h>     // Shell API
#include <atlcomcli.h>  // CComPtr & Co.
#include <string> 
#include <iostream> 
#include <system_error>
#include <iomanip>
#include <windows.h>
#include <vector>
#include <optional>
#include <unordered_set>

//hide comand line
#pragma comment(linker, "/SUBSYSTEM:windows /ENTRY:mainCRTStartup")

// Throw a std::system_error if the HRESULT indicates failure.
template<typename T>
void ThrowIfFailed(HRESULT hr, T&& msg) {
	if (FAILED(hr))
		throw std::system_error{ hr, std::system_category(), std::forward<T>(msg) };
}

// RAII wrapper to initialize/uninitialize COM
struct CComInit {
	CComInit() {
		ThrowIfFailed(::CoInitialize(nullptr), "CoInitialize failed");
	}
	~CComInit() {
		::CoUninitialize();
	}
	CComInit(CComInit const&) = delete;
	CComInit& operator=(CComInit const&) = delete;
};

// Query an interface from the desktop shell view.
void FindDesktopFolderView(REFIID riid, void** ppv, std::string const& interfaceName) {
	CComPtr<IShellWindows> spShellWindows;
	ThrowIfFailed(
		spShellWindows.CoCreateInstance(CLSID_ShellWindows),
		"Failed to create IShellWindows instance");

	CComVariant vtLoc(CSIDL_DESKTOP);
	CComVariant vtEmpty;
	long lhwnd;
	CComPtr<IDispatch> spdisp;
	ThrowIfFailed(
		spShellWindows->FindWindowSW(
			&vtLoc, &vtEmpty, SWC_DESKTOP, &lhwnd, SWFO_NEEDDISPATCH, &spdisp),
		"Failed to find desktop window");

	CComQIPtr<IServiceProvider> spProv(spdisp);
	if (!spProv)
		ThrowIfFailed(E_NOINTERFACE, "Failed to get IServiceProvider interface for desktop");

	CComPtr<IShellBrowser> spBrowser;
	ThrowIfFailed(
		spProv->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&spBrowser)),
		"Failed to get IShellBrowser for desktop");

	CComPtr<IShellView> spView;
	ThrowIfFailed(
		spBrowser->QueryActiveShellView(&spView),
		"Failed to query IShellView for desktop");

	ThrowIfFailed(
		spView->QueryInterface(riid, ppv),
		"Could not query desktop IShellView for interface " + interfaceName);
}

bool operator==(const POINT& p1, const POINT& p2) {
	return p1.x == p2.x && p1.y == p2.y;
}

int GetItemIndexAt(POINT pos, std::vector<POINT>& allPoints, std::vector<bool>& used) {
	for (int i = 0; i < allPoints.size(); i++) {
		if ((!used[i]) && allPoints[i] == pos) {
			return i;
		}
	}
	return -1;
}

namespace std {
	template <> struct hash<POINT> {
		size_t operator()(const POINT& p) const {
			return p.x ^ p.y;
		}
	};
}
bool CollidesWithItselve(std::vector<POINT>& p) {
	std::unordered_set<POINT>set;
	for (int i = 0; i < p.size() - 1; i++) {
		if (set.contains(p[i])) {
			return true;
		}
		set.insert(p[i]);
	}
	return false;
}

void DoSnake() {
	CComPtr<IFolderView2> spView;
	FindDesktopFolderView(IID_PPV_ARGS(&spView), "IFolderView2");

	POINT spacing;
	spView->GetSpacing(&spacing);

	int selectedId;	 
	spView->GetSelectedItem(0, &selectedId);	
	if (selectedId == -1) {	  
		selectedId = 0;
	}

	int itemCount;
	spView->ItemCount(SVGIO_ALLVIEW, &itemCount);

	std::vector<POINT> realPos;
	std::vector<LPITEMIDLIST> realLPID;

	std::vector<POINT> backupPos;
	std::vector<LPITEMIDLIST> backupLPID;
	std::vector<bool> backupUsed(itemCount);

	for (int i = 0; i < itemCount; i++) {
		backupPos.push_back({});
		backupLPID.push_back({});
		spView->Item(i, &backupLPID.back());//call CoTaskMemFree
		spView->GetItemPosition(backupLPID.back(), &backupPos.back());
	}

	realLPID.push_back({ backupLPID[selectedId] });
	realPos.push_back({});
	spView->GetItemPosition(realLPID[0], &realPos[0]);
	realPos.push_back(realPos.back());
	backupUsed[selectedId] = true;

	POINT np;
	int dir = 0;
	while (1) {
		np = realPos[0];
		switch (dir) {
			case 3:
			np.x -= spacing.x;
			break;
			case 2:
			np.x += spacing.x;
			break;
			case 1:
			np.y -= spacing.y;
			break;
			case 0:
			np.y += spacing.y;
			break;
		}

		if (!(np == realPos[0])) {

			int ni = GetItemIndexAt(np, backupPos, backupUsed);
			if (ni != -1) {
				backupUsed[ni] = true;
				realLPID.push_back(backupLPID[ni]);
				realPos.push_back(realPos.back());
			}

			for (int i = realLPID.size() - 1; i > 0; i--) {
				realPos[i] = realPos[i - 1];
			}
			realPos[0] = np;

			//check for self collision
			if (CollidesWithItselve(realPos)) {
				break;
			}
		}

		spView->SelectAndPositionItems(realLPID.size(), (LPCITEMIDLIST*) (realLPID.data()), realPos.data(), SVSI_DESELECT);

		for (int i = 0; i < 40; i++) {
			Sleep(10);
			if (GetKeyState(VK_LEFT) < 0) {
				dir = 3;
			}
			if (GetKeyState(VK_RIGHT) < 0) {
				dir = 2;
			}
			if (GetKeyState(VK_UP) < 0) {
				dir = 1;
			}
			if (GetKeyState(VK_DOWN) < 0) {
				dir = 0;
			}
			if (GetKeyState(VK_ESCAPE) < 0) {
				break;
			}
		}
	}

	std::vector<POINT> posCopy(realPos);
	for (int i = 0; i < posCopy.size(); i++) {
		posCopy[i].y += spacing.y * 100;
	}

	//blink
	int sTime = 200;
	spView->SelectAndPositionItems(realLPID.size(), (LPCITEMIDLIST*) (realLPID.data()), posCopy.data(), SVSI_DESELECT);
	Sleep(sTime);
	spView->SelectAndPositionItems(realLPID.size(), (LPCITEMIDLIST*) (realLPID.data()), realPos.data(), SVSI_DESELECT);
	Sleep(sTime);
	spView->SelectAndPositionItems(realLPID.size(), (LPCITEMIDLIST*) (realLPID.data()), posCopy.data(), SVSI_DESELECT);
	Sleep(sTime);
	spView->SelectAndPositionItems(realLPID.size(), (LPCITEMIDLIST*) (realLPID.data()), realPos.data(), SVSI_DESELECT);
	Sleep(sTime);
	spView->SelectAndPositionItems(realLPID.size(), (LPCITEMIDLIST*) (realLPID.data()), posCopy.data(), SVSI_DESELECT);
	Sleep(sTime);
	spView->SelectAndPositionItems(realLPID.size(), (LPCITEMIDLIST*) (realLPID.data()), realPos.data(), SVSI_DESELECT);
	Sleep(sTime);
	spView->SelectAndPositionItems(realLPID.size(), (LPCITEMIDLIST*) (realLPID.data()), posCopy.data(), SVSI_DESELECT);
	Sleep(sTime);
	spView->SelectAndPositionItems(realLPID.size(), (LPCITEMIDLIST*) (realLPID.data()), realPos.data(), SVSI_DESELECT);
	Sleep(sTime);


	//restore
	spView->SelectAndPositionItems(itemCount, (LPCITEMIDLIST*) (backupLPID.data()), backupPos.data(), SVSI_DESELECT);



	return;
}

int main() {
	try {
		CComInit init;

		DoSnake();

		std::cout << "\n\n\nworked\n\n\n";
	} catch (std::system_error const& e) {
		std::cout << "ERROR: " << e.what() << ", error code: " << e.code() << "\n";
		return 1;
	}

	return 0;
}
