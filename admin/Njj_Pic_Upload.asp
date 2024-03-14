<!--#include file="admin_check.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet"  href="style.css" type="text/css">
<title>图片上传</title>
<script language="javascript">
//图片预览
function viewmypic(mypic,file) {
        if (file.value){
             mypic.src=file.value;
             mypic.style.display="";
             mypic.border=1;
             pricesum.innerHTML = "现在,请点击上传按钮";
          }

}

//图片格式检查
function checkImage(sId)
{
  if(( document.all[sId].value.indexOf(".gif") == 1) && (document.all[sId].value.indexOf(".jpg") == 1)&& (document.all[sId].value.indexOf(".bmp") == 1) && ( document.all[sId].value.indexOf(".jpeg") == 1) && (document.all[sId].value.indexOf(".png") == 1) && (document.all[sId].value.indexOf(".swf") == 1)) {
    alert("请先点击浏览按钮选择gif或jpg或jpeg或png或bmp格式的文件");
    event.returnValue = false;
    }
    else
    {
    esave.style.visibility="visible";
}
}

var imgObj;
function checkImg(theURL,winName){
  // 对象是否已创建
  if (typeof(imgObj) == "object"){
    // 是否已取得了图像的高度和宽度
    if ((imgObj.width != 0) && (imgObj.height != 0))
      // 根据取得的图像高度和宽度设置弹出窗口的高度与宽度，并打开该窗口
      // 其中的增量 20 和 30 是设置的窗口边框与图片间的间隔量
      OpenFullSizeWindow(theURL,winName, ",width=" + (imgObj.width+50) + ",height=" + (imgObj.height+80));
    else
      // 因为通过 Image 对象动态装载图片，不可能立即得到图片的宽度和高度，所以每隔100毫秒重复调用检查
      setTimeout("checkImg('" + theURL + "','" + winName + "')", 100)
  }
}

function OpenFullSizeWindow(theURL,winName,features) {
  var aNewWin, sBaseCmd;
  // 弹出窗口外观参数
  sBaseCmd = "toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=no,";
  // 调用是否来自 checkImg 
  if (features == null || features == ""){
    // 创建图像对象
    imgObj = new Image();
    // 设置图像源
    imgObj.src = theURL;
    // 开始获取图像大小
    checkImg(theURL, winName)
  }
  else{
    // 打开窗口
    var x=imgObj.width+50;
    if (x<400)
      {
       x=400;
      }
    aNewWin=window.resizeTo(x, imgObj.height+120); 

    // 聚焦窗口
    aNewWin.focus();

  }
}

</script>
</head>

<body onload="javascript:pricesum.innerHTML = '请点击浏览按钮,选择您要上传的图片文件';">
<form name="form1" method="post" action="njj_Pic_UpFile.asp" enctype="multipart/form-data" onsubmit="checkImage('FormName')">
    <center>
	<img name="showimg" id="showimg" src="" height=300 width=300 style="display:none;" alt="预览图片"><br><font color="#FF0000"><span id="pricesum"></span><br>
    <input type="file" name="FormName" value="" id="file" onchange="viewmypic(showimg,this.form.file);OpenFullSizeWindow(document.all.showimg.src,'','');return false" size="27">
    <input type="hidden" name="filepath" value="../uploadpic/">
    <input type="hidden" name="action" value="<%=request("action")%>">
    <input type="hidden" name="Fname" value="<%=request("Fname")%>">
    <input type="hidden" name="flag" value="<%=request("flag")%>"><input type="submit" name="Submit" value="上传"></font></center>
  </form>

</body>

</html>