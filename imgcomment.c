/*
 * ImgComment.exe — Excel 图片批注工具 (v7 final)
 *
 * Compile:
 *   x86_64-w64-mingw32-gcc -o ImgComment.exe imgcomment_v7.c icon_res.o \
 *       -lole32 -loleaut32 -luuid -lgdi32 -lgdiplus -lcomdlg32 -lshell32 \
 *       -lcomctl32 -static-libgcc -Os -s -mwindows
 */

#define WIN32_LEAN_AND_MEAN
#define _WIN32_IE 0x0600
#include <windows.h>
#include <commctrl.h>
#include <oleauto.h>
#include <commdlg.h>
#include <gdiplus.h>
#include <shellapi.h>
#include <stdio.h>
#include <stdlib.h>

#define C_BG      RGB(240, 243, 248)
#define C_LINK    RGB(40, 100, 230)

static HINSTANCE g_hInst = NULL;
static HFONT g_fnt = NULL;
static HCURSOR g_hand = NULL;

/* ============================================================== */
/*  Crash handler                                                  */
/* ============================================================== */
static LONG WINAPI crash_handler(EXCEPTION_POINTERS *ep)
{
    WCHAR buf[512];
    swprintf(buf, 512,
        L"ImgComment 发生意外错误\nUnexpected error\n\n"
        L"Exception: 0x%08X\nAddress: 0x%p\n\n"
        L"请截图发给开发者。\nScreenshot & send to developer.",
        ep->ExceptionRecord->ExceptionCode,
        ep->ExceptionRecord->ExceptionAddress);
    MessageBoxW(NULL, buf, L"ImgComment · 崩溃 / Crash", MB_OK|MB_ICONERROR);
    return EXCEPTION_EXECUTE_HANDLER;
}

/* ============================================================== */
/*  COM                                                           */
/* ============================================================== */
static IDispatch *g_xl = NULL;
static ULONG_PTR g_gp = 0;

static VARIANT vb(const WCHAR *s){VARIANT v;VariantInit(&v);v.vt=VT_BSTR;v.bstrVal=SysAllocString(s);return v;}
static VARIANT vi(int i){VARIANT v;VariantInit(&v);v.vt=VT_I4;v.lVal=i;return v;}
static VARIANT vd(double d){VARIANT v;VariantInit(&v);v.vt=VT_R8;v.dblVal=d;return v;}

static HRESULT xc(IDispatch *o,WORD f,VARIANT *r,const WCHAR *n,UINT c,VARIANT *a){
    if(!o)return E_POINTER;DISPID id;WCHAR *ns[]={(WCHAR*)n};
    HRESULT h=o->lpVtbl->GetIDsOfNames(o,&IID_NULL,ns,1,LOCALE_USER_DEFAULT,&id);
    if(FAILED(h))return h;
    DISPPARAMS p={a,NULL,c,0};DISPID pd=DISPID_PROPERTYPUT;
    if(f==DISPATCH_PROPERTYPUT){p.cNamedArgs=1;p.rgdispidNamedArgs=&pd;}
    EXCEPINFO ex={0};return o->lpVtbl->Invoke(o,id,&IID_NULL,LOCALE_USER_DEFAULT,f,&p,r,&ex,NULL);
}
#define XG(o,p,r) xc(o,DISPATCH_PROPERTYGET,r,p,0,NULL)
#define XP(o,p,a) xc(o,DISPATCH_PROPERTYPUT,NULL,p,1,a)
#define XM(o,m)   xc(o,DISPATCH_METHOD,NULL,m,0,NULL)

static int gpi(void){if(g_gp)return 1;GdiplusStartupInput in={1};return GdiplusStartup(&g_gp,&in,NULL)==Ok?1:0;}
static int gsz(const WCHAR *p,UINT *w,UINT *h){
    if(!gpi())return 0;GpImage *im=NULL;
    if(GdipLoadImageFromFile(p,&im)!=Ok)return 0;
    GdipGetImageWidth(im,w);GdipGetImageHeight(im,h);GdipDisposeImage(im);return 1;
}

static void msg(const WCHAR *t,const WCHAR *s){MessageBoxW(NULL,t,s,MB_OK|MB_ICONWARNING);}

static int xci(void){
    if(g_xl)return 1;
    CLSID c;if(FAILED(CLSIDFromProgID(L"Excel.Application",&c))){msg(L"无法找到 Excel\nFailed to find Excel.",L"错误 / Error");return 0;}
    IUnknown *u=NULL;
    if(SUCCEEDED(GetActiveObject(&c,NULL,&u))&&u){
        HRESULT h=u->lpVtbl->QueryInterface(u,&IID_IDispatch,(void**)&g_xl);u->lpVtbl->Release(u);
        if(SUCCEEDED(h)&&g_xl)return 1;
    }
    if(FAILED(CoCreateInstance(&c,NULL,CLSCTX_LOCAL_SERVER,&IID_IDispatch,(void**)&g_xl))){
        msg(L"无法启动 Excel。请先打开 Excel 再试。\nCannot start Excel.",L"错误 / Error");return 0;
    }
    VARIANT vt;VariantInit(&vt);vt.vt=VT_BOOL;vt.boolVal=VARIANT_TRUE;XP(g_xl,L"Visible",&vt);
    vt.boolVal=VARIANT_FALSE;XP(g_xl,L"DisplayAlerts",&vt);return 1;
}

/* ============================================================== */
/*  Insert (with full NULL checks)                                */
/* ============================================================== */
static int x_ins(const WCHAR *path){
    if(!xci())return 0;
    VARIANT c;VariantInit(&c);XG(g_xl,L"ActiveCell",&c);
    if(c.vt!=VT_DISPATCH||!c.pdispVal){msg(L"请先在 Excel 中选中一个单元格\nPlease select a cell first.",L"提示 / Notice");return 0;}
    VARIANT m;VariantInit(&m);XG(c.pdispVal,L"Comment",&m);
    if(m.vt==VT_DISPATCH&&m.pdispVal){XM(m.pdispVal,L"Delete");m.pdispVal->lpVtbl->Release(m.pdispVal);}
    VariantInit(&m);xc(c.pdispVal,DISPATCH_METHOD,&m,L"AddComment",0,NULL);
    if(m.vt!=VT_DISPATCH||!m.pdispVal){if(c.pdispVal)c.pdispVal->lpVtbl->Release(c.pdispVal);msg(L"添加批注失败。请检查是否关闭「新式批注」\nAddComment failed.",L"错误 / Error");return 0;}
    VARIANT e=vb(L"");XP(m.pdispVal,L"Text",&e);VariantClear(&e);
    VARIANT s;VariantInit(&s);XG(m.pdispVal,L"Shape",&s);
    if(s.vt==VT_DISPATCH&&s.pdispVal){
        UINT iw=0,ih=0;if(gsz(path,&iw,&ih)&&iw>0&&ih>0){VARIANT vw=vd(iw*0.75);XP(s.pdispVal,L"Width",&vw);VARIANT vh=vd(ih*0.75);XP(s.pdispVal,L"Height",&vh);}
        VARIANT vk=vi(-1);XP(s.pdispVal,L"LockAspectRatio",&vk);
        VARIANT f;VariantInit(&f);XG(s.pdispVal,L"Fill",&f);
        if(f.vt==VT_DISPATCH&&f.pdispVal){VARIANT vt=vi(-1);XP(f.pdispVal,L"Visible",&vt);VARIANT p=vb(path);xc(f.pdispVal,DISPATCH_METHOD,NULL,L"UserPicture",1,&p);VariantClear(&p);f.pdispVal->lpVtbl->Release(f.pdispVal);}
        s.pdispVal->lpVtbl->Release(s.pdispVal);
    }
    if(m.pdispVal)m.pdispVal->lpVtbl->Release(m.pdispVal);
    if(c.pdispVal)c.pdispVal->lpVtbl->Release(c.pdispVal);
    return 1;
}

/* ============================================================== */
/*  Paste                                                         */
/* ============================================================== */
static int x_paste(void){
    if(!OpenClipboard(NULL)){msg(L"无法打开剪贴板\nCannot open clipboard.",L"错误 / Error");return 0;}
    HBITMAP h=(HBITMAP)GetClipboardData(CF_BITMAP);
    if(!h){CloseClipboard();msg(L"剪贴板中没有图片。请先截图或复制。\nNo image in clipboard.",L"提示 / Notice");return 0;}
    CloseClipboard();
    WCHAR t[MAX_PATH];GetTempPathW(MAX_PATH,t);wcscat(t,L"ib.bmp");
    BITMAP bm;GetObject(h,sizeof(bm),&bm);HDC dc=GetDC(NULL);
    if(!dc){msg(L"获取屏幕 DC 失败",L"错误");return 0;}
    BITMAPINFO bi={0};bi.bmiHeader.biSize=sizeof(BITMAPINFOHEADER);bi.bmiHeader.biWidth=bm.bmWidth;bi.bmiHeader.biHeight=-bm.bmHeight;
    bi.bmiHeader.biPlanes=1;bi.bmiHeader.biBitCount=24;bi.bmiHeader.biCompression=BI_RGB;
    DWORD r=((bm.bmWidth*24+31)/32)*4,tb=r*bm.bmHeight;void *d=malloc(tb);
    if(!d){ReleaseDC(NULL,dc);msg(L"内存不足\nOut of memory.",L"错误");return 0;}
    GetDIBits(dc,h,0,bm.bmHeight,d,&bi,DIB_RGB_COLORS);
    BITMAPFILEHEADER bf={0};bf.bfType=0x4D42;bf.bfSize=sizeof(bf)+sizeof(BITMAPINFOHEADER)+tb;bf.bfOffBits=sizeof(bf)+sizeof(BITMAPINFOHEADER);
    HANDLE hf=CreateFileW(t,GENERIC_WRITE,0,NULL,CREATE_ALWAYS,FILE_ATTRIBUTE_NORMAL,NULL);
    if(hf!=INVALID_HANDLE_VALUE){DWORD w;WriteFile(hf,&bf,sizeof(bf),&w,NULL);WriteFile(hf,&bi.bmiHeader,sizeof(BITMAPINFOHEADER),&w,NULL);WriteFile(hf,d,tb,&w,NULL);CloseHandle(hf);}
    else{free(d);ReleaseDC(NULL,dc);msg(L"无法写入临时文件",L"错误");return 0;}
    free(d);ReleaseDC(NULL,dc);int r2=x_ins(t);DeleteFileW(t);return r2;
}

static void x_del(void){
    if(!xci())return;
    VARIANT c;VariantInit(&c);XG(g_xl,L"ActiveCell",&c);
    if(c.vt!=VT_DISPATCH||!c.pdispVal){msg(L"请先在 Excel 中选中一个单元格\nPlease select a cell first.",L"提示 / Notice");return;}
    VARIANT m;VariantInit(&m);XG(c.pdispVal,L"Comment",&m);
    if(m.vt==VT_DISPATCH&&m.pdispVal){XM(m.pdispVal,L"Delete");m.pdispVal->lpVtbl->Release(m.pdispVal);}
    else{msg(L"当前单元格没有批注\nNo comment found.",L"提示 / Notice");if(c.pdispVal)c.pdispVal->lpVtbl->Release(c.pdispVal);return;}
    if(c.pdispVal)c.pdispVal->lpVtbl->Release(c.pdispVal);
}

/* ============================================================== */
/*  Help                                                          */
/* ============================================================== */
static void show_help(void){
    MessageBoxW(NULL,
        L"📖 ImgComment 使用帮助\n\n"
        L"1️⃣ 选中 Excel 单元格\n"
        L"2️⃣ 点击「插入」选择图片\n"
        L"   或「粘贴」从剪贴板粘贴\n"
        L"3️⃣ 点击「删除」清除批注\n\n"
        L"⚠ Excel→文件→选项→常规→取消「新式批注」\n\n"
        L"官网: https://mc.mcgg.cc",
        L"ImgComment · 帮助",MB_OK|MB_ICONINFORMATION);
}

/* ============================================================== */
/*  Owner-draw helpers                                            */
/* ============================================================== */
static void draw_round_rect(HDC dc,RECT *r,COLORREF clr,int rad){
    HPEN pn=CreatePen(PS_SOLID,0,clr),op=SelectObject(dc,pn);
    HBRUSH br=CreateSolidBrush(clr),ob=SelectObject(dc,br);
    RoundRect(dc,r->left,r->top,r->right,r->bottom,rad*2,rad*2);
    SelectObject(dc,ob);DeleteObject(br);SelectObject(dc,op);DeleteObject(pn);
}

/* Simple font (system default, bold) */
static void mkfnt(void){
    NONCLIENTMETRICSW nm={sizeof(nm)};SystemParametersInfoW(SPI_GETNONCLIENTMETRICS,sizeof(nm),&nm,0);
    nm.lfMessageFont.lfHeight=-13;nm.lfMessageFont.lfWeight=FW_BOLD;nm.lfMessageFont.lfQuality=CLEARTYPE_QUALITY;
    g_fnt=CreateFontIndirectW(&nm.lfMessageFont);g_hand=LoadCursorW(NULL,(LPCWSTR)IDC_HAND);
}

#define BN_INS 301
#define BN_PST 302
#define BN_DEL 303
#define BN_INF 304
#define BN_EXT 305
#define BN_LNK 306

static const WCHAR *B[]={L"📷  插入 / Insert",L"📋  粘贴 / Paste",L"🗑  删除 / Delete",L"ℹ   帮助 / Help",L"✕   退出 / Exit"};
static int    BID[]={BN_INS,BN_PST,BN_DEL,BN_INF,BN_EXT};
static COLORREF BCLR[]={RGB(75,130,220),RGB(75,130,220),RGB(200,80,80),RGB(130,140,155),RGB(170,120,110)};
static COLORREF BHVR[]={RGB(95,150,240),RGB(95,150,240),RGB(220,100,100),RGB(150,160,175),RGB(190,140,130)};
#define BH 32
#define BG 6

/* ============================================================== */
/*  WndProc                                                       */
/* ============================================================== */
static LRESULT CALLBACK WndProc(HWND h,UINT m,WPARAM w,LPARAM l)
{
    static int hover=-1;
    switch(m){
    case WM_CREATE:{
        mkfnt();HWND p=h;int y=10;
        for(int i=0;i<5;i++){
            CreateWindowW(L"BUTTON",B[i],WS_VISIBLE|WS_CHILD|BS_OWNERDRAW,10,y,200,BH,p,(HMENU)(INT_PTR)BID[i],g_hInst,NULL);
            if(g_fnt)SendMessageW(GetDlgItem(p,BID[i]),WM_SETFONT,(WPARAM)g_fnt,TRUE);y+=BH+BG;
        }
        y+=2;CreateWindowW(L"STATIC",L"",WS_VISIBLE|WS_CHILD|SS_ETCHEDHORZ,10,y,200,2,p,NULL,NULL,NULL);y+=8;
        HWND hl=CreateWindowW(L"STATIC",L"https://mc.mcgg.cc",WS_VISIBLE|WS_CHILD|SS_NOTIFY|SS_CENTER,10,y,200,18,p,(HMENU)(INT_PTR)BN_LNK,g_hInst,NULL);
        if(g_fnt)SendMessageW(hl,WM_SETFONT,(WPARAM)g_fnt,TRUE);return 0;
    }
    case WM_CTLCOLORSTATIC:{
        HDC dc=(HDC)w;
        if(GetDlgCtrlID((HWND)l)==BN_LNK){SetTextColor(dc,C_LINK);SetBkColor(dc,C_BG);return (LRESULT)GetStockObject(DC_BRUSH);}
        SetBkColor(dc,C_BG);return (LRESULT)GetStockObject(DC_BRUSH);
    }
    case WM_SETCURSOR:{
        HWND cw=(HWND)w;
        if(cw==GetDlgItem(h,BN_LNK)){SetCursor(g_hand);SetWindowLongPtrW(h,DWLP_MSGRESULT,TRUE);return TRUE;}
        return DefWindowProcW(h,m,w,l);
    }
    case WM_DRAWITEM:{
        LPDRAWITEMSTRUCT is=(LPDRAWITEMSTRUCT)l;int idx=-1;
        for(int j=0;j<5;j++)if(BID[j]==is->CtlID){idx=j;break;}
        if(idx<0)return TRUE;
        BOOL hv=(hover==is->CtlID);COLORREF bg=hv?BHVR[idx]:BCLR[idx];
        draw_round_rect(is->hDC,&is->rcItem,bg,4);
        SetBkMode(is->hDC,TRANSPARENT);SetTextColor(is->hDC,RGB(255,255,255));
        RECT tr=is->rcItem;tr.left+=8;
        DrawTextW(is->hDC,B[idx],(int)wcslen(B[idx]),&tr,DT_VCENTER|DT_SINGLELINE|DT_LEFT);return TRUE;
    }
    case WM_MOUSEMOVE:{
        POINT pt={LOWORD(l),HIWORD(l)};HWND c=ChildWindowFromPoint(h,pt);int nh=-1;
        if(c){int id=GetDlgCtrlID(c);for(int j=0;j<5;j++)if(BID[j]==id){nh=id;break;}}
        if(nh!=hover){hover=nh;for(int j=0;j<5;j++){HWND b=GetDlgItem(h,BID[j]);if(b)InvalidateRect(b,NULL,TRUE);}}
        return 0;
    }
    case WM_COMMAND:
        switch(LOWORD(w)){
        case BN_INS:{
            OPENFILENAMEW of={sizeof(of),0};WCHAR p[MAX_PATH]={0};
            of.lpstrFilter=L"Images (*.jpg;*.png;*.bmp)\0*.jpg;*.jpeg;*.png;*.bmp\0All\0*.*\0";
            of.lpstrFile=p;of.nMaxFile=MAX_PATH;of.lpstrTitle=L"选择图片 / Select Image";
            of.Flags=OFN_FILEMUSTEXIST|OFN_HIDEREADONLY;
            if(GetOpenFileNameW(&of))x_ins(p);return 0;
        }
        case BN_PST:x_paste();return 0;
        case BN_DEL:x_del();return 0;
        case BN_INF:show_help();return 0;
        case BN_LNK:ShellExecuteW(NULL,L"open",L"https://mc.mcgg.cc",NULL,NULL,SW_SHOW);return 0;
        case BN_EXT:PostQuitMessage(0);return 0;
        }
        return 0;
    case WM_DESTROY:case WM_CLOSE:PostQuitMessage(0);return 0;
    }
    return DefWindowProcW(h,m,w,l);
}

/* ============================================================== */
/*  Entry                                                         */
/* ============================================================== */
int WINAPI WinMain(HINSTANCE h,HINSTANCE,LPSTR,int s)
{
    SetUnhandledExceptionFilter(crash_handler);
    g_hInst=h;CoInitializeEx(NULL,COINIT_APARTMENTTHREADED);InitCommonControls();
    WNDCLASSW wc={0};wc.lpfnWndProc=WndProc;wc.hInstance=h;
    wc.hIcon=LoadIconW(h,MAKEINTRESOURCEW(1));
    wc.hCursor=LoadCursorW(NULL,(LPCWSTR)IDC_ARROW);
    wc.hbrBackground=CreateSolidBrush(C_BG);wc.lpszClassName=L"ICV11";
    if(!RegisterClassW(&wc)){CoUninitialize();return 1;}
    int W=220,H=10+5*(BH+BG)+2+8+22;RECT r={0,0,W,H};
    AdjustWindowRectEx(&r,WS_OVERLAPPED|WS_CAPTION|WS_SYSMENU|WS_MINIMIZEBOX,FALSE,WS_EX_TOPMOST);
    HWND hw=CreateWindowExW(WS_EX_TOPMOST,L"ICV11",L"ImgComment · Excel 图片批注工具",
        WS_OVERLAPPED|WS_CAPTION|WS_SYSMENU|WS_VISIBLE|WS_MINIMIZEBOX,
        CW_USEDEFAULT,CW_USEDEFAULT,r.right-r.left,r.bottom-r.top,NULL,NULL,h,NULL);
    if(!hw){CoUninitialize();return 1;}
    SetWindowPos(hw,HWND_TOPMOST,0,0,0,0,SWP_NOMOVE|SWP_NOSIZE);ShowWindow(hw,s);
    MSG msg;while(GetMessageW(&msg,NULL,0,0)){TranslateMessage(&msg);DispatchMessageW(&msg);}
    if(g_xl){g_xl->lpVtbl->Release(g_xl);}
    if(g_gp)GdiplusShutdown(g_gp);if(g_fnt)DeleteObject(g_fnt);
    CoUninitialize();return 0;
}
