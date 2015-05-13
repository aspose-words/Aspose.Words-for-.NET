var success = "success";
var error = "error";

$(window).load(function () {
    initializeTooltips();
});

// Initialize the tool tips on default page
function initializeTooltips() {
    // Download document link
    $('.download-document').tooltip({
        'show': true,
        'placement': 'top',
        'title': "Download the document"
    });

    // Delete document link
    $('.delete-document').tooltip({
        'show': true,
        'placement': 'top',
        'title': "Delete the document"
    });

    // Upload document link
    $('.upload-document').tooltip({
        'show': true,
        'placement': 'top',
        'title': "Upload document"
    });

    // Create new folder link
    $('.create-folder').tooltip({
        'show': true,
        'placement': 'top',
        'title': "Create new folder"
    });
}

function btnCreateFolder_onClick() {
    var lblCurrentFolder = $(".lblCurrentFolder").text().replace(/\\/g, "\\\\");
    var txtCreateFolder = $("#txtCreateFolder").val();

    // Call web method to create the folder
    $.ajax({
        type: "POST",
        url: "Default.aspx/CreateFolder",
        data: '{currentFolder: "' + lblCurrentFolder + '" , folderName: "' + txtCreateFolder + '" }',
        contentType: "application/json; charset=utf-8",
        dataType: "json",
        success: function (data) {
            if (data.d == success) {
                $("#CreateFolderDialog").modal("hide");
                $("form").submit();
            }
            else {
                // Something wrong. Display the error message in alert
                // Remove the hidden class and add the alert-danger class
                $("#divAlert").removeClass("hidden");
                $("#divAlert").addClass("alert-danger");
                $("#divAlert").text(data.d);
            }
        },
        failure: function (data) {
            alert('error');
        }
    });
}

// Reset the alert messages and warnings when dialog is closed
$('#CreateFolderDialog').on('hidden.bs.modal', function (e) {
    $("#divAlert").addClass("hidden");
    $("#divAlert").removeClass("alert-danger");
    $("#divAlert").removeClass("alert-info");
    $("#divAlert").removeClass("alert-success");
    $("#divAlert").removeClass("alert-warning");

    $("#txtCreateFolder").val("");
})

$(".delete-confirm").confirm({
    text: "Are you sure you want to delete?",
    title: "Delete",
    //confirm: function(button) {
    //    //delete();
    //},
    cancel: function (button) {
        // nothing to do
    },
    confirmButton: "Yes",
    cancelButton: "No",
    post: true,
    confirmButtonClass: "btn-danger",
    cancelButtonClass: "btn-default"
});

function viewDocument(filePath, isFolder, added, deleted) {
    resetDocumentViewerModalData();
    // If it is a folder, dont do anything, let server handle it
    if (isFolder == "True") {
        return true;
    }
    else {
        // Get file name from path
        var fileNameIndex = filePath.lastIndexOf("\\") + 1;
        var fileName = filePath.substr(fileNameIndex);
        // Show the document viewer modal dialog
        $("#DocumentViewerDialog").modal();
        // Set title
        $("#DocumentViewerDialogTitle").text(fileName);
        getDocumentData(filePath);

        // Update the summary, if available (in case of comparison)
        if (added != null && deleted != null)
        {
            $("#DocumentViewerSummary").removeClass("hidden");
            $("#DocumentViewerSummaryAdded").text(added);
            $("#DocumentViewerSummaryDeleted").text(deleted);
        }
        return false;
    }
}

function getDocumentData(filePath) {
    filePath = filePath.replace(/\\/g, "\\\\");
    $.ajax({
        type: "POST",
        url: "Default.aspx/GetDocumentData",
        data: '{ filePath: "' + filePath + '" , sessionID: "' + $("#txtSessionID").val() + '" }',
        contentType: "application/json; charset=utf-8",
        dataType: "json",
        success: function (data) {
            // If there is error
            if (data.d[0].substr(0, 5) == error) {
                $("#DocumentViewerAlert").addClass("alert-danger");
                $("#DocumentViewerAlert").removeClass("hidden");
                $("#DocumentViewerAlert").text(data.d);
            }
            else {

                // In case call is successful, pass data to success method
                getDocumentData_Success(data.d);
            }
        },
        failure: function (data) {
            alert('error');
        }
    });
}

$('#chSelectAll').click(function () {
    var c = this.checked;
    //alert("hello " + c + " - " + $(".select-document").size());
    $(".select-document :checkbox").prop('checked', c);
});

function getDocumentData_Success(result) {
    var totalPages = result[1];
    var imageFolder = result[2];
    //alert(totalPages);
    // Show the first page
    $("#CurrentDocumentPage").attr("src", imageFolder + "/0.png");
    // Show pagination
    $("#DocumentViewerPagination").removeClass("hidden");
    // Add pages in pagination
    for (var iPage = 1 ; iPage <= totalPages ; iPage++) {
        var currentPage = 'setCurrentPage(&quot;' + imageFolder + '/' + (iPage - 1) + '.png' + '&quot;)';
        //alert(currentPage);
        $("#DocumentViewerPaginationUL li:nth-child(" + iPage + ")")
            .after('<li class="DocumentViewerPaginationLI"><a onclick="' + currentPage + '" href="#">' + iPage + '</a></li>');
    }
}

function resetDocumentViewerModalData() {
    // Reset the alert
    $("#DocumentViewerAlert").addClass("hidden");
    $("#DocumentViewerAlert").removeClass("alert-danger");
    $("#DocumentViewerAlert").removeClass("alert-info");
    $("#DocumentViewerAlert").removeClass("alert-success");
    $("#DocumentViewerAlert").removeClass("alert-warning");

    // Hide the pagination
    $("#DocumentViewerPagination").addClass("hidden");
    // Remove all the pagination LI items with class DocumentViewerPaginationLI
    $(".DocumentViewerPaginationLI").remove();

    // Reset the default blank image for current page
    $("#CurrentDocumentPage").attr("src", "http://localhost:50465/Temp/temp.png");

    // Hide the summary
    $("#DocumentViewerSummary").addClass("hidden");
}

function setCurrentPage(currentPage) {
    $("#CurrentDocumentPage").attr("src", currentPage);
}

//$(".select-document").on('click', function (event) {
//    var checkboxes = $('.select-document :checked');
//    // If more than 2 documents are selected, show error and return
//    if (checkboxes.length > 2)
//    {
//        $("#PageGeneralDialog").modal();
//        $("#PageGeneralDivAlert").removeClass("hidden");
//        $("#PageGeneralDivAlert").addClass("alert-danger");
//        $("#PageGeneralDivAlert").text("Please select only 2 documents for comparison.");
//        return false;
//    }
//});

// Compare two selected documents
function btnCompare_onClick() {
    var checkboxes = $('.select-document :checked');
    // If more than 2 documents are selected, show error and return
    if (checkboxes.length != 2) {
        $("#PageGeneralDialog").modal();
        $("#PageGeneralDivAlert").removeClass("hidden");
        $("#PageGeneralDivAlert").addClass("alert-danger");
        $("#PageGeneralDivAlert").text("Please select 2 documents for comparison.");
        return false;
    }
    
    var documents = new Array();
    checkboxes.each(function (index, elem) {
        var documentName = $(elem).parent().parent().parent().find(".link-document").val();
        $("#divPageAlert").append(documentName + " , ");
        documents.push(documentName);
    });

    // Replace the \ with \\, it is special character
    documents[0] = documents[0].replace(/\\/g, "\\\\");
    documents[1] = documents[1].replace(/\\/g, "\\\\");

    compareDocuments(documents[0], documents[1]);
}

// Compare two selected documents
function btnCompareURLs_onClick() {
    var documents = new Array();
    documents[0] = $("#txtFirstURL").val();
    documents[1] = $("#txtSecondURL").val();
    //alert("hello");

    compareDocuments(documents[0], documents[1]);
}

// This is generic method that will take URL of two documents for comparison (WEB or File)
function compareDocuments(document1, document2) {
    resetDocumentViewerModalData();
    // Call server side method to compare the documents
    $.ajax({
        type: "POST",
        url: "Default.aspx/CompareDocuments",
        data: '{ document1: "' + document1 + '" , document2: "' + document2 + '" }',
        contentType: "application/json; charset=utf-8",
        dataType: "json",
        success: function (data) {
            // If there is error
            if (data.d[0].substr(0, 5) == error) {
                $("#DocumentViewerAlert").addClass("alert-danger");
                $("#DocumentViewerAlert").removeClass("hidden");
                $("#DocumentViewerAlert").text(data.d);
                $("#DocumentViewerDialog").modal();
            }
            else {

                // In case call is successful, pass data to success method
                var comparisonDocument = data.d[1];
                viewDocument(comparisonDocument, "False", data.d[2], data.d[3]);
            }
        },
        failure: function (data) {
            alert('error');
        }
    });
}