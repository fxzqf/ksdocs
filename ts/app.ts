window.onload=()=>{
    let url=location.href.split("?")[1];
    let parames=new URLSearchParams(url);
    alert(parames.keys.length);
}