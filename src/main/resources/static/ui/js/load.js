$(document).ready(function(){

});

function postData(){
    var fileName = getFileName($("#file").val());
    if(fileName==null||fileName==""){
        alert("请选择上传文件");
        return;
    }
    var formData = new FormData();
    formData.append("file",$("#file")[0].files[0]);
    $.ajax({
        url:'/excel/receive',
        type:'post',
        data: formData,
        contentType: false,
        processData: false,
        success:function(res){
            console.log(res);
            if(res.data["code"]=="succ"){
                alert('成功');
            }else if(res.data["code"]=="err"){
                alert('失败');
            }else{
                console.log(res);
            }
        }
    })
}

function getFileName(o){
    var pos=o.lastIndexOf("\\");
    return o.substring(pos+1);
}