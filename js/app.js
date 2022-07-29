"use strict";
let taskPanes = new Array();
//let app = wps.EtApplication().Application;
function OnAddinLoad(ribbonUI) {
    wps.RibbonUI = ribbonUI;
    return true;
}
function openBook(obj) {
    //wps.PluginStorage.getItem()
    //let App=wps.EtApplication().Application;
    //App.Workbooks.Add();
    //aap.Visible=true;
}
function OnAction(control) {
}
function OnGetEnabled(control) {
    return true;
}
/**
* 获取一个控件的图标
* @param control 要获取图标的控件
* @returns 图标的SVG图像的URL
*/
function GetImage(control) {
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
function onWorkbookOpen(wb) {
}
/**
 *
 * @param wb
 */
function onWorkbookBeforeClose(wb) {
}
/**
 * 当窗口不活动时隐藏工作簿对应的操作窗格
 */
function onWindowDeactivate(wb, win) {
    taskPanes.forEach(element => {
        if (wb.FullName == element.wb.FullName)
            element.tp.Visible = false;
    });
}
/**
 * 当窗口激活时显示工作簿对应的操作窗格
 * @param wb
 * @param win
 * @returns
 */
function onWindowActivate(wb, win) {
    taskPanes.forEach(element => {
        if (wb.FullName == element.wb.FullName)
            element.tp.Visible = true;
    });
}
/**
 *
 */
window.onload = () => {
    if (wps.Application)
        wps.Application = wps.EtApplication();
    if (wps.Application.Workbooks.Count == 0) {
        wps.Application.Workbooks.Add();
        wps.Application.Visible = true;
        let tp1 = wps.CreateTaskPane("https://fxzqf.github.io/wpsapp/about.html", "关于表格助手");
        taskPanes.push({ wb: wps.Application.ActiveWorkbook, tp: tp1 });
        tp1.Visible = true;
    }
    else {
        for (let i = 1; i <= wps.Application.Workbooks.Count; i++) {
        }
    }
    wps.ApiEvent.AddApiEventListener("WindowActivate", onWindowActivate);
    wps.ApiEvent.AddApiEventListener("WindowDeactivate", onWindowDeactivate);
    wps.ApiEvent.AddApiEventListener("WorkbookBeforeClose", onWorkbookBeforeClose);
    wps.ApiEvent.AddApiEventListener("WorkbookOpen", onWorkbookOpen);
};
