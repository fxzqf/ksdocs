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
    taskPanes.forEach(element => {
        if ((wb as Et.EtWorkbook).FullName == element.wb.FullName) {
            taskPanes.splice(taskPanes.indexOf(element), 1);
            element.tp.Delete();
        }
    });
}
/**
 * 当窗口不活动时隐藏工作簿对应的操作窗格
 */
function onWindowDeactivate(wb: object, win: object) {
    taskPanes.forEach(element => {
        if ((<Et.EtWorkbook>wb).FullName == element.wb.FullName) element.tp.Visible = false;
    });

}

/**
 * 当窗口激活时显示工作簿对应的操作窗格
 * @param wb
 * @param win
 * @returns
 */
function onWindowActivate(wb: object, win: object) {
    taskPanes.forEach(element => {
        if ((wb as Et.EtWorkbook).FullName == element.wb.FullName) element.tp.Visible = true;
    });
}

/**
 * 
 */
window.onload = () => {

    if (wps.Application) wps.Application = wps.EtApplication();
    if (wps.Application.Workbooks.Count == 0) {
        wps.Application.Workbooks.Add();
        wps.Application.Visible = true;
        let tp1 = wps.CreateTaskPane("https://fxzqf.github.io/wpsapp/about.html", "关于表格助手");
        taskPanes.push({ wb: wps.Application.ActiveWorkbook, tp: tp1 });
        tp1.Visible = true;
    }
    else {
        for (let i = 1; i <= wps.Application.Workbooks.Count; i++) {
            var wb1 = (<Et.EtWorkbook>wps.Application.Workbooks.Item(i));
            var obj = wb1.CustomDocumentProperties;
            for (var x = obj.Count; x > 0; x--) {
                if (obj.Item(x).Name == "TaskPane") {
                    taskPanes.push({ wb: wb1, tp: wps.CreateTaskPane("https://fxzqf.github.io/" + obj.Item(x).Value, "表格助手") });
                }
            }
        }
    }
    wps.ApiEvent.AddApiEventListener("WindowActivate", onWindowActivate);
    wps.ApiEvent.AddApiEventListener("WindowDeactivate", onWindowDeactivate);
    wps.ApiEvent.AddApiEventListener("WorkbookBeforeClose", onWorkbookBeforeClose);
    wps.ApiEvent.AddApiEventListener("WorkbookOpen", onWorkbookOpen);
}
