// main.cpp
#include <windows.h>
#include <tchar.h>
#define ID_MYTIMER 100
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
ATOM InitApp(HINSTANCE);
BOOL InitInstance(HINSTANCE, int);
void MyDrawText(HWND, HDC,PAINTSTRUCT);

char szClassName[] = "sample02";	//ウィンドウクラス

int WINAPI WinMain(HINSTANCE hCurInst, HINSTANCE hPrevInst, LPSTR lpsCmdLine, int nCmdShow) {
	MSG msg;
	BOOL bRet;

	if (!InitApp(hCurInst))
		return FALSE;
	if (!InitInstance(hCurInst, nCmdShow))
		return FALSE;

	while ((bRet = GetMessage(&msg, NULL, 0, 0)) != 0) {
		if (bRet == -1) {
			break;
		}
		else {
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
	return (int)msg.wParam;
}

ATOM InitApp(HINSTANCE hInst) {
	WNDCLASSEX wc;
	wc.cbSize = sizeof(WNDCLASSEX);
	wc.style = CS_HREDRAW | CS_VREDRAW;
	wc.lpfnWndProc = WndProc;	//プロシージャ名
	wc.cbClsExtra = 0;
	wc.cbWndExtra = 0;
	wc.hInstance = hInst;//インスタンス
	wc.hIcon = NULL; //アプリのアイコン。.icoファイルをリソースファイルに読み込みここに記入
	wc.hCursor = (HCURSOR)LoadImage(NULL, MAKEINTRESOURCE(IDC_ARROW), IMAGE_CURSOR, 0, 0, LR_DEFAULTSIZE | LR_SHARED);
	//wc.hbrBackground = (HBRUSH)GetStockObject(BLACK_BRUSH);//Windowの背景色を指定。
	wc.hbrBackground = (HBRUSH)GetStockObject(WHITE_BRUSH); //透過用のWindow背景色を設定
	wc.lpszMenuName = NULL;	 // メニュー名。リソースファイルで設定した値を記入
	wc.lpszClassName = (LPCSTR)szClassName;
	wc.hIconSm = NULL; //アプリのアイコンの小さい版。タスクバーに表示されるもの

	return (RegisterClassEx(&wc));
}


//ウィンドウの生成
BOOL InitInstance(HINSTANCE hInst, int nCmdShow) {
	HWND hWnd;

	//hWnd = CreateWindow(szClassName,
	hWnd = CreateWindowEx(WS_EX_LAYERED, //ウィンドウにレイヤードウィンドウ拡張属性を付加する
		szClassName,
		"Sample Application 1", //Windowのタイトル
		WS_OVERLAPPEDWINDOW, //ウィンドウの種類
		CW_USEDEFAULT,	//Ｘ座標 (指定なしはCW_USEDEFAULT)
		CW_USEDEFAULT,	//Ｙ座標(指定なしはCW_USEDEFAULT)
		CW_USEDEFAULT,	//幅(指定なしはCW_USEDEFAULT)
		CW_USEDEFAULT,	//高さ(指定なしはCW_USEDEFAULT)
		NULL, //親ウィンドウのハンドル、親を作るときはNULL
		NULL, //メニューハンドル、クラスメニューを使うときはNULL
		hInst, //インスタンスハンドル
		NULL);

	if (!hWnd)
		return FALSE;
	ShowWindow(hWnd, nCmdShow);
	UpdateWindow(hWnd);
	return TRUE;
}


//ウィンドウプロシージャ
//ユーザーがWindowを操作したり、Windowが作られたりした場合、この関数が呼び出されて処理を行う。
LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wp, LPARAM lp) {
	HBRUSH hBrush;
	HDC hdc, hdc_mem;
	PAINTSTRUCT ps;
	static HPEN hBigPen, hSmallPen;
	RECT rc;
	int id,w,h,i,recx,recy;
	static int x,y,wx, wy, r;
	char szBuf[32] = "猫でもわかるLayer";
	static BOOL bDec = FALSE;
	switch (msg) {
	case WM_CREATE:
		//CreateWindow()でWindowを作成したときに呼び出される。ここでウィジェットを作成する。
		/*
		BOOL SetLayeredWindowAttributes(
			HWND hwnd,           // ウィンドウのハンドル
			COLORREF crKey,      // COLORREF値
			BYTE bAlpha,         // アルファの値
			DWORD dwFlags        // アクションフラグ
		);
		hwndには、レイヤーウィンドウのハンドルを指定します。
		crKeyには、透明にするカラーキーのRGB値を指定します。
		bAlphaには、アルファ値を指定します。0で全くの透明、255で不透明となります。
		dwFlagsには、次の値のいずれか、または両方を指定します。
		値				意味
		LWA_COLORKEY	crKeyが有効
		LWA_ALPHA		bAlphaが有効
		成功すると0以外の値が返り、失敗すると0が返ります。
		*/
		hBigPen = CreatePen(PS_SOLID, 4, RGB(0, 0, 0));
		hSmallPen = CreatePen(PS_SOLID, 1, RGB(0, 0, 0));
		SetLayeredWindowAttributes(hWnd, RGB(255, 0, 0), 0,LWA_COLORKEY); //透過色として赤(255,0,0)を指定
		//タイマー動作開始
		//SetTimer(hWnd, ID_MYTIMER, 200, NULL);
		break;
	case WM_TIMER:
		//Timerイベント
		if (wp != ID_MYTIMER) {
			return DefWindowProc(hWnd, msg, wp, lp);
		}
		if (bDec) {
			r -= 5;
			if (r <= 0) {
				bDec = FALSE;
			}
		}
		else {
			r += 5;
			if (r >= wy || r >= wx) {
				bDec = TRUE;
			}
		}
		InvalidateRect(hWnd, NULL, TRUE);	//これによりWM_PAINTメッセージが再発行される
		break;
	case WM_DESTROY:
		//ユーザーがWindow右上の×ボタンを押すとここが実行される
		DeleteObject(hBigPen);
		DeleteObject(hSmallPen);
		KillTimer(hWnd, ID_MYTIMER);
		PostQuitMessage(0); //終了メッセージ
		break;
	case WM_SIZE:
		//サイズ変更メッセージ
		wx = LOWORD(lp);
		wy = HIWORD(lp);
		x = wx / 2;
		y = wy / 2;
		break;
	case WM_PAINT:
		//画面に図形などを描く処理を実装
		hdc = BeginPaint(hWnd, &ps);	//デバイスコンテキスト取得
		hBrush = CreateSolidBrush(RGB(255, 0, 0));	//Brushのハンドル取得
		SelectObject(hdc, hBrush);	//ブラシオブジェクトを適用
		/*
		BOOL ExtFloodFill(
		HDC hdc,          // デバイスコンテキストハンドル
		int nXStart,      // 開始点の x 座標
		int nYStart,      // 開始点の y 座標
		COLORREF crColor, // 色
		UINT fuFillType   // 種類
		);
		現在のブラシを使って塗りつぶします。
		hdcには、デバイスコンテキストハンドルを指定します。
		nXStart, nYStartには、塗りつぶし開始の座標を指定します。
		crColorは、RGB値を指定しますが、その意味はfuFillTypeにより異なります。
		fuFillTypeには、塗りつぶしの種類を指定します。次のいずれかを指定します。
		値	意味
		FLOODFILLBORDER	crColorで指定した色が囲んでいる領域を、塗りつぶします。
		FLOODFILLSURFACE	crColorで指定した色と同じ色になっている領域を、塗りつぶします。
		*/
		ExtFloodFill(hdc, 1, 1, RGB(255, 255, 255), FLOODFILLSURFACE);	//最初に白(255,255,255)で塗りつぶしているため、そこをターゲットにする
		/*
		//図形描画してみる碁盤の目
		GetClientRect(hWnd, &rc);
		w = rc.right;
		h = rc.bottom;
		SelectObject(hdc, hBigPen);
		Rectangle(hdc, 30, 30, w - 30, h - 30);
		SelectObject(hdc, hSmallPen);
		for (i = 0;i < 10;i++) {
			recy = 30 + ((h - 60) / 10)*i;
			MoveToEx(hdc, 30, recy, NULL);
			LineTo(hdc, w - 30, recy);
			recx = 30 + ((w - 60) / 10)*i;
			MoveToEx(hdc, recx, 30, NULL);
			LineTo(hdc, recx, h - 30);
		}
		//円を描画
		Ellipse(hdc, x - (r / 2), y - (r / 2), x + (r / 2), y + (r / 2));
		*/
		//テキスト描画してみる
		//TextOut(hdc, 10, 90, szBuf, (int)strlen(szBuf));
		MyDrawText(hWnd, hdc,ps);
		DeleteObject(hBrush);
		EndPaint(hWnd, &ps);
		break;
	default:
		return (DefWindowProc(hWnd, msg, wp, lp));
	}
	return 0;
}
//拡張テキスト表示プロシージャ
//Return void
//args
//HWND hWnd		ウィンドウのインスタンスハンドル
//HDC hdc		デバイスコンテキストのハンドル
//PAINTSTRUCT ps	PAINTSTRUCT構造体
void MyDrawText(HWND hWnd, HDC hdc,PAINTSTRUCT ps) {
	char szSTR[]= _T("5P8A1314P003A");
	HFONT hFont1 = CreateFont();
	SelectObject(hdc, hFont1);
	TextOut(hdc, 10, 90, (LPCSTR)szSTR, (int)strlen(szSTR));
}