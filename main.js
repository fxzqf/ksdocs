function newWorkBook()
{
    App=wps.EtApplication().Application;
    App.Workbooks.Add();
    App.CreateTaskPane("https://zhibiao.uicp.fun/").Visible=true
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
            return "images/1.svg";
        case "btnShowDialog":
            return "images/2.svg";
        case "btnShowTaskPane":
            return "images/3.svg";
        default:
            ;
    }
    return "images/newFromTemp.svg";
}
