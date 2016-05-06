$(function () {
    var jcrop_api, zoomRation,ratio;
    $preview = $('#preview-pane');
    $pcnt = $('#preview-pane .preview-container');
    $pimg = $('#preview-pane .preview-container img');
    xsize = $pcnt.width();
    ysize = $pcnt.height();
    var boundx;
    var boundy;
    var defaultWidth = 500; //默认宽度

    $('#fileupload').fileupload({
        maxNumberOfFiles : 1,
        dataType: 'json',
        // 上传完成后的执行逻辑
        done: function (e, data) {
            //图片原始高度
            var height = data.result.height;
            //图片原始宽度
            var width = data.result.width;

            //图片的长宽比
            ratio = height / width;

            //图片缩放比
            zoomRation = width / defaultWidth;

            /*对图片在界面上的显示进行缩放*/
            $("#PicBox").attr("width", defaultWidth);   
            $("#PicBox").attr("height", defaultWidth * ratio);

            
            $("#PicBox").attr("src", "/UploadFiles/" + data.result.fileName);

            $pimg.attr("src", "/UploadFiles/" + data.result.fileName);

            $("#imgUrl").val("/UploadFiles/" + data.result.fileName);
            initJcrop();
        },
        add: function (e, data) {
            $("#preview-pane").addClass("hidden");
            if (jcrop_api!=undefined) {
                jcrop_api.destroy();
            }            
            data.submit();
        }
    }).prop('disabled', !$.support.fileInput)
        .parent().addClass($.support.fileInput ? undefined : 'disabled');

    function initJcrop() {        
        $('#PicBox').Jcrop({
            aspectRatio: 1 / 1,
            onRelease: releaseCheck,
            onChange: updatePreview,
            onSelect: updatePreview,
            bgColor: "black",
            bgOpacity: ".4",
            setSelect: [ 100, 100, 330, 330 ]
        }, function () {
            jcrop_api = this;
            var bounds = this.getBounds();
            boundx = bounds[0];
            boundy = bounds[1];
            $preview.appendTo(jcrop_api.ui.holder);
            $("#preview-pane").removeClass("hidden");
        });        
    }

    function updatePreview(c) {
        if (parseInt(c.w) > 0) {
            var rx = xsize / c.w;
            var ry = ysize / c.h;
            $pimg.css({
                width: Math.round(rx * boundx) + 'px',
                height: Math.round(rx * boundy) + 'px',
                marginLeft: '-' + Math.round(rx * c.x) + 'px',
                marginTop: '-' + Math.round(ry * c.y) + 'px'
            });

            $("#x").val(c.x * zoomRation);
            $("#y").val(c.y * zoomRation);
            $("#w").val(c.w * zoomRation);
            $("#h").val(c.h * zoomRation);
        }
    };


    function updateCoords(c) {
        $("#btnCropAndSave").removeClass("hidden");
        $('#x').val(Math.round(c.x * zoomRation));
        $('#y').val(Math.round(c.y * zoomRation));
        $('#w').val(Math.round(c.w * zoomRation));
        $('#h').val(Math.round(c.h * zoomRation));
        $('#img').val($("#PicBox").attr("src"));
    };

    function releaseCheck() {
        jcrop_api.setOptions({ allowSelect: true });
    };
});