function exportMaterial() {
    $('.ajax-loader').show();
    var url = "/Home/ExportMaterialExcel";
    PostData(url, {}, exportMaterialResponse, null);
}
function exportMaterialResponse() { window.location =  "/Home/ExportMaterialExcel"; }
function PostData(url, _data, _successHandler, ShowBlackImage) {


    if (ShowBlackImage == null || ShowBlackImage == undefined) {
        ShowBlackImage = true;
    }
    $.ajax({
        type: 'POST',
        url: url,
        data: _data,
        success: _successHandler,
        global: ShowBlackImage
    });
}