$(document).ready(function(){

});

function uploadExcel(){
    var fileName = getFileName($("#file").val());
    if(fileName==null||fileName==""){
        alert("请选择上传文件");
        return;
    }
    var isTotal = $("#isTotal").prop('checked');
    var colName = $("#colName").val();
    //var idCol = $("#idCol").val();
    var valueCol = $("#valueCol").val();
    var formData = new FormData();
    var flag = false;
    $(".fileList span").each(function(){
        var isTotal = $(this).attr("isTotal");
        if(isTotal = "true"){
            flag = true;
        }
    })
    if(flag==true&&isTotal=="true"){
        alert("已上传了汇总表");
        return;
    }
    console.log(colName);
    formData.append("file",$("#file")[0].files[0]);
    formData.append("isTotal",isTotal);
    formData.append("colName",colName);
    /*formData.append("idCol",idCol);*/
    formData.append("valueCol",valueCol);
    $.ajax({
        url:'/excel/receive',
        type:'post',
        data: formData,
        contentType: false,
        processData: false,
        success:function(data){
            if(data.success=="true"){
                alert("上传成功");

                var fileName = "<span isTotal='"+data.isTotal+"'>"+data.fileName+"</span>";

                var html = $(".fileList").html();
                if(html==""){
                    $(".fileList").append(fileName);
                }else{
                    $(".fileList").append(","+fileName);
                }


            }
        }
    })
}

function clean(){
    $.ajax({
        url:'/excel/clean',
        type:'post',
        data: {},
        contentType: false,
        processData: false,
        success:function(data){
            if(data.success=="true"){
                alert("清空成功");
                $(".fileList").html("");
            }
        }
    })
}

function create(){
    window.open('/excel/create');
}

function getFileName(o){
    var pos=o.lastIndexOf("\\");
    return o.substring(pos+1);
}