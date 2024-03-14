var imgObj;
function checkImg(theURL,winName){
  // 对象是否已创建
  if (typeof(imgObj) == "object"){
    // 是否已取得了图像的高度和宽度
    if ((imgObj.width != 0) && (imgObj.height != 0))
      // 根据取得的图像高度和宽度设置弹出窗口的高度与宽度，并打开该窗口
      // 其中的增量 20 和 30 是设置的窗口边框与图片间的间隔量
      OpenFullSizeWindow(theURL,winName, ",width=" + (imgObj.width+20) + ",height=" + (imgObj.height+30));
    else
      // 因为通过 Image 对象动态装载图片，不可能立即得到图片的宽度和高度，所以每隔100毫秒重复调用检查
      setTimeout("checkImg('" + theURL + "','" + winName + "')", 100)
  }
}

function OpenFullSizeWindow(theURL,winName,features) {
  var aNewWin, sBaseCmd;
  // 弹出窗口外观参数
  sBaseCmd = "toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no,";
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
    aNewWin = window.open(theURL,winName, sBaseCmd + features);
    // 聚焦窗口
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

