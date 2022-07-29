"use strict";
/// <reference path="index.d.ts" />
declare namespace wps {
    export let RibbonUI: Kso.KsoRibbonUI;
    export let Enum: any;
    export let Application: Et.EtApplication;
    export let ActiveTaskPane: wps.CustomTaskpane | null;
}
let taskPanes: Array<{ wb: Et.EtWorkbook, tp: wps.CustomTaskpane }> = new Array();

//let app = wps.EtApplication().Application;
function OnAddinLoad(ribbonUI: Kso.KsoRibbonUI) {

    wps.RibbonUI = ribbonUI;

    return true;
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
function onWorkbookOpen(wb: object) {

}
/**
 * 
 * @param wb 
 */
function onWorkbookBeforeClose(wb: object) {

}
/**
 * 
 */
function onWindowDeactivate(b: object, win: object) {

}

/**
 * 当窗口激活时显示工作簿对应的操作窗格
 * @param wb
 * @param win
 * @returns
 */
function onWindowActivate(wb: object, win: object) {
    taskPanes.forEach(element => {


    });


    //for (var _i = 0, taskPanes_2 = taskPanes; _i < taskPanes_2.length; _i++) {


    //    var taskPane = taskPanes_2[_i];
    //    if (taskPane.Name == name) {
    //        (wps.ActiveTaskPane = wps.GetTaskPane(taskPane.ID)).Visible = true;
    //        return;
    //    }
    //}
}

/**
 * 
 */
window.onload = () => {

    if (wps.Application) wps.Application = wps.EtApplication();
    if (wps.Application.Workbooks.Count == 0) {
        wps.Application.Workbooks.Add();
        wps.Application.Visible = true;
        let tp1= wps.CreateTaskPane("https://fxzqf.github.io/wpsapp/about.html", "欢迎使用表格助手");
        tp1.Visible = true;
        let wb1=wps.Application.ActiveWorkbook;
        taskPanes.push({wb:wb1,tp:tp1});
    }
    wps.ApiEvent.AddApiEventListener("WindowActivate", onWindowActivate);
    wps.ApiEvent.AddApiEventListener("WindowDeactivate", onWindowDeactivate);
    wps.ApiEvent.AddApiEventListener("WorkbookBeforeClose", onWorkbookBeforeClose);
    wps.ApiEvent.AddApiEventListener("WorkbookOpen", onWorkbookOpen);
}