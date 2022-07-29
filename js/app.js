"use strict";
let taskPanes = new Array();
//let app = wps.EtApplication().Application;
function OnAddinLoad(ribbonUI) {
    if (wps.Application.Workbooks.Count == 0)
        wps.Application.Workbooks.Add();
    wps.RibbonUI = ribbonUI;
    //wps.Application.WindowState=Et.EtXlWindowState.xlMaximized;
    wps.Application.Visible = true;
    wps.CreateTaskPane("https://zhibiao.uicp.fun/", "表格助手").Visible = true;
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
 *
 */
function onWindowDeactivate(b, win) {
}
/**
 * 当窗口激活时显示工作簿对应的操作窗格
 * @param wb
 * @param win
 * @returns
 */
function onWindowActivate(wb, win) {
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
    if (wps.Application)
        wps.Application = wps.EtApplication();
    wps.ApiEvent.AddApiEventListener("WindowActivate", onWindowActivate);
    wps.ApiEvent.AddApiEventListener("WindowDeactivate", onWindowDeactivate);
    wps.ApiEvent.AddApiEventListener("WorkbookBeforeClose", onWorkbookBeforeClose);
    wps.ApiEvent.AddApiEventListener("WorkbookOpen", onWorkbookOpen);
};
