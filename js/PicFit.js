var imgObj;
function checkImg(theURL,winName){
  // �����Ƿ��Ѵ���
  if (typeof(imgObj) == "object"){
    // �Ƿ���ȡ����ͼ��ĸ߶ȺͿ��
    if ((imgObj.width != 0) && (imgObj.height != 0))
      // ����ȡ�õ�ͼ��߶ȺͿ�����õ������ڵĸ߶����ȣ����򿪸ô���
      // ���е����� 20 �� 30 �����õĴ��ڱ߿���ͼƬ��ļ����
      OpenFullSizeWindow(theURL,winName, ",width=" + (imgObj.width+20) + ",height=" + (imgObj.height+30));
    else
      // ��Ϊͨ�� Image ����̬װ��ͼƬ�������������õ�ͼƬ�Ŀ�Ⱥ͸߶ȣ�����ÿ��100�����ظ����ü��
      setTimeout("checkImg('" + theURL + "','" + winName + "')", 100)
  }
}

function OpenFullSizeWindow(theURL,winName,features) {
  var aNewWin, sBaseCmd;
  // ����������۲���
  sBaseCmd = "toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no,";
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
    aNewWin = window.open(theURL,winName, sBaseCmd + features);
    // �۽�����
    aNewWin.focus();
  }
}

var imageObject;
function ResizeImage(obj, MaxW, MaxH)
{
    if (obj != null) imageObject = obj;
    var state=imageObject.readyState;
    var oldImage = new Image();
    oldImage.src = imageObject.src;
    var dW=oldImage.width; var dH=oldImage.height;
    if(dW>MaxW || dH>MaxH) {
        a=dW/MaxW; b=dH/MaxH;
        if(b>a) a=b;
        dW=dW/a; dH=dH/a;
    }
    if(dW > 0 && dH > 0)
        imageObject.width=dW;imageObject.height=dH;
    if(state!='complete' || imageObject.width>MaxW || imageObject.height>MaxH) {
        setTimeout("ResizeImage(null,"+MaxW+","+MaxH+")",40);
    }
}

function ImagePreload() { 
	var args = ImagePreload.arguments;
	document.ImgPreArray = new Array(args.length);
	for(var i=0; i<args.length; i++) {
		document.ImgPreArray[i] = new Image;
		document.ImgPreArray[i].src = "uploadpic/"+ args[i];
	}
}

function findItem(n, d) {
	var p,x,i;
	if(!d) d=document;
	if((p=n.indexOf("?"))>0&&parent.frames.length) {
		d=parent.frames[n.substring(p+1)].document;
		n=n.substring(0,p);
	}
	if(!(x=d[n])&&d.all)
		x=d.all[n];
	for (i=0;!x&&i<d.forms.length;i++)
		x=d.forms[i][n];
	for(i=0;!x&&d.layers&&i<d.layers.length;i++)
		x=findItem(n,d.layers[i].document);
	return x;
}

function fitSize() {
	var a, b;
	var imgobj = document.all("ShowImage");
	var oldimg = new Image();
	oldimg.src = imgobj.src;
	var dw = oldimg.width;
	var dh = oldimg.height;
	if(imgobj == null) {
		setTimeout("fitSize()", 50);
		return;
	}
	if(imgobj.offsetWidth == 0) {
		setTimeout("fitSize()", 50);
		return;
	}
	var maxW = 220;
	var maxH = 220;
	if(dw>maxW || dh>maxH) {
		a = dw/maxW;
		b = dh/maxW; 
		if(b>a) a=b;
		dw = dw/a;
		dh = dh/a;
	}
	if(dw > 0 && dh > 0) {
		imgobj.width = dw;
		imgobj.height = dh;
	}
}

function GetShowImg(imgtext, imgfile) {
	document.all("ShowImgText").innerHTML = imgtext;
	document.all("ShowImage").src = "uploadpic/"+ imgfile;
	fitSize();
}

function showlist(dd)
{
  if(dd=="a")
  {
   linkimg.style.display="";
  }
  else
  {   
   linkimg.style.display="none";
  }
}

