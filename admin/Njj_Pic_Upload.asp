<!--#include file="admin_check.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet"  href="style.css" type="text/css">
<title>ͼƬ�ϴ�</title>
<script language="javascript">
//ͼƬԤ��
function viewmypic(mypic,file) {
        if (file.value){
             mypic.src=file.value;
             mypic.style.display="";
             mypic.border=1;
             pricesum.innerHTML = "����,�����ϴ���ť";
          }

}

//ͼƬ��ʽ���
function checkImage(sId)
{
  if(( document.all[sId].value.indexOf(".gif") == 1) && (document.all[sId].value.indexOf(".jpg") == 1)&& (document.all[sId].value.indexOf(".bmp") == 1) && ( document.all[sId].value.indexOf(".jpeg") == 1) && (document.all[sId].value.indexOf(".png") == 1) && (document.all[sId].value.indexOf(".swf") == 1)) {
    alert("���ȵ�������ťѡ��gif��jpg��jpeg��png��bmp��ʽ���ļ�");
    event.returnValue = false;
    }
    else
    {
    esave.style.visibility="visible";
}
}

var imgObj;
function checkImg(theURL,winName){
  // �����Ƿ��Ѵ���
  if (typeof(imgObj) == "object"){
    // �Ƿ���ȡ����ͼ��ĸ߶ȺͿ��
    if ((imgObj.width != 0) && (imgObj.height != 0))
      // ����ȡ�õ�ͼ��߶ȺͿ�����õ������ڵĸ߶����ȣ����򿪸ô���
      // ���е����� 20 �� 30 �����õĴ��ڱ߿���ͼƬ��ļ����
      OpenFullSizeWindow(theURL,winName, ",width=" + (imgObj.width+50) + ",height=" + (imgObj.height+80));
    else
      // ��Ϊͨ�� Image ����̬װ��ͼƬ�������������õ�ͼƬ�Ŀ�Ⱥ͸߶ȣ�����ÿ��100�����ظ����ü��
      setTimeout("checkImg('" + theURL + "','" + winName + "')", 100)
  }
}

function OpenFullSizeWindow(theURL,winName,features) {
  var aNewWin, sBaseCmd;
  // ����������۲���
  sBaseCmd = "toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=no,";
  // �����Ƿ����� checkImg 
  if (features == null || features == ""){
    // ����ͼ�����
    imgObj = new Image();
    // ����ͼ��Դ
    imgObj.src = theURL;
    // ��ʼ��ȡͼ���С
    checkImg(theURL, winName)
  }
  else{
    // �򿪴���
    var x=imgObj.width+50;
    if (x<400)
      {
       x=400;
      }
    aNewWin=window.resizeTo(x, imgObj.height+120); 

    // �۽�����
    aNewWin.focus();

  }
}

</script>
</head>

<body onload="javascript:pricesum.innerHTML = '���������ť,ѡ����Ҫ�ϴ���ͼƬ�ļ�';">
<form name="form1" method="post" action="njj_Pic_UpFile.asp" enctype="multipart/form-data" onsubmit="checkImage('FormName')">
    <center>
	<img name="showimg" id="showimg" src="" height=300 width=300 style="display:none;" alt="Ԥ��ͼƬ"><br><font color="#FF0000"><span id="pricesum"></span><br>
    <input type="file" name="FormName" value="" id="file" onchange="viewmypic(showimg,this.form.file);OpenFullSizeWindow(document.all.showimg.src,'','');return false" size="27">
    <input type="hidden" name="filepath" value="../uploadpic/">
    <input type="hidden" name="action" value="<%=request("action")%>">
    <input type="hidden" name="Fname" value="<%=request("Fname")%>">
    <input type="hidden" name="flag" value="<%=request("flag")%>"><input type="submit" name="Submit" value="�ϴ�"></font></center>
  </form>

</body>

</html>