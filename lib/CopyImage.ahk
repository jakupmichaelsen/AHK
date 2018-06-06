#Include Gdip_All.ahk

CopyImage(ImageFile) {
    pToken := Gdip_Startup()
    Gdip_SetBitmapToClipboard(pBitmap := Gdip_CreateBitmapFromFile(ImageFile))
    Gdip_DisposeImage(pBitmap)
    Gdip_Shutdown(pToken)
}