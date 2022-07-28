"use strict";
/// <reference path="index.d.ts" />
declare namespace wps {
    export let RibbonUI: Kso.KsoRibbonUI;
    export let Enum: any;
    export let Application: Et.EtApplication;
    export let ActiveTaskPane: wps.CustomTaskpane | null;
}


let app = wps.EtApplication().Application;
function OnAddinLoad(ribbonUI: Kso.KsoRibbonUI) {
    alert("Addin");
    if (app.Workbooks.Count == 0) app.Workbooks.Add();
    wps.RibbonUI = ribbonUI;
    //app.WindowState=Et.EtXlWindowState.xlMaximized;
    app.Visible = true;
    wps.CreateTaskPane("https://zhibiao.uicp.fun/", "表格助手").Visible = true;
    return true;
}
window.onload=()=>{
    alert("Window");
}
function openBook(obj: string) {
    //wps.PluginStorage.getItem()

    //let App=wps.EtApplication().Application;
    //App.Workbooks.Add();





    //aap.Visible=true;
}
function OnAction(control: Kso.KsoRibbonControl) {

}
function OnGetEnabled(control: Kso.KsoRibbonControl) {
    return true;
}
/**
* 获取一个控件的图标
* @param control 要获取图标的控件
* @returns 图标的SVG图像的URL
*/
function GetImage(control: any) {
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