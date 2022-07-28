/// <reference path="./src/index.d.ts" />
let app=wps.EtApplication().Application;
function OnAddinLoad(ribbonUI:any) {
    if(app.Workbooks.Count==0) app.Workbooks.Add();
    wps.CreateTaskPane("https://zhibiao.uicp.fun/","表格助手").Visible=true;
    //app.Visible=true;
    return true;
}
function openBook(obj:string)
{
    //wps.PluginStorage.getItem()
    
    //let App=wps.EtApplication().Application;
    //App.Workbooks.Add();

    

    
    
    //aap.Visible=true;
}
function OnAction(control:any)
{
    
}
function OnGetEnabled(control:any)
{
    return true;
}
/**
* 获取一个控件的图标
* @param control 要获取图标的控件
* @returns 图标的SVG图像的URL
*/
function GetImage(control:any) {
    var eleId = control.Id;
    switch (eleId) {
        case "btnShowMsg":
            return "./images/1.svg";
        case "btnShowDialog":
            return "./images/2.svg";
        case "btnShowTaskPane":
            return "./images/3.svg";
        default:
        ;
    }
return "./images/newFromTemp.svg";
}