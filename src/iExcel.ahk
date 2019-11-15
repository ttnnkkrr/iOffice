

class iExcel {
    
    obj := ""
    oWkbk := ""
    oSheet := ""

    __New(saveas = false, visible = true, sheetID = 1){
        if !IsObject(this.GetWorkBook(saveas))
            this.New(saveas, visible)
        this.sheet(sheetID)
    }
    ;; returns workbook object
    New(saveas = false, visible = true){
        this.obj := ComObjCreate("excel.application")
        this.obj.visible := visible
        this.oWkbk := this.obj.workbooks.add
        if (saveas && this.__validateFilePath(saveas)) {
           this.obj.DisplayAlerts := False
           this.oWkbk.saveas(saveas)
           this.obj.DisplayAlerts := True
        }
        return this.oWkbk
    }

    ;; get an existing workbook. returns false if a valid file isnt referenced
    GetWorkBook(path = false){
        ControlGet, hwnd, hwnd, , Excel71, ahk_class XLMAIN
        this.oWkbk := false
        if hwnd {
            window := this.__ObjectFromWindow(hwnd,-16)
            if (path && this.__validateFilePath(path)){
                if (path == window.parent.FullName){
                    try this.oWkbk := window.application.workbooks.open(path)
                }
                else{
                    this.oWkbk := ComObjGet(path)
                }
            }
            else{
                this.oWkbk := window.parent
            }
        }
        else{
            if (path && this.__validateFilePath(path)){
                this.oWkbk := ComObjGet(path)
            }            
        }
        return this.oWkbk
    }

    ;; void
    Sheet(ID = false){
        if ID
            this.oSheet := this.oWkbk.sheets(ID)
        this.oSheet.activate
    }


    ;; void
    Range(range = false, value = false){
        if IsObject(this.oSheet) && range {
            if value
                this.oSheet.range(range).value := value
            return this.oSheet.range(range).value
        } 
        else { ;; something didnt get setup correctly
            Throw, "Please set up a workbook"
        }
    }
    
    ;;https://www.codeproject.com/Tips/216238/Regular-Expression-to-Validate-File-Path-and-Exten
    ;; validate a path is valid
    __validateFilePath(path){
        return RegExMatch(path, "^(?:[\w]\:|\\)(\\[a-z_\-\s0-9\.]+)+\.(txt|xls|xlsx|csv)$") && FileExist(path)
    }

    ;***borrowd & tweaked from Acc.ahk Standard Library*** by Sean  Updated by jethrow*****************
    __ObjectFromWindow(hWnd, idObject = -4){ 
        (if Not h)?h:=DllCall("LoadLibrary","Str","oleacc","Ptr")
            If DllCall("oleacc\AccessibleObjectFromWindow","Ptr",hWnd,"UInt",idObject&=0xFFFFFFFF,"Ptr",-VarSetCapacity(IID,16)+NumPut(idObject==0xFFFFFFF0?0x46000000000000C0:0x719B3800AA000C81,NumPut(idObject==0xFFFFFFF0?0x0000000000020400:0x11CF3C3D618736E0,IID,"Int64"),"Int64"), "Ptr*", pacc)=0
                Return	ComObjEnwrap(9,pacc,1)
    }
}

