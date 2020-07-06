(function () {
    'use strict';

    function main($scope, $http, Upload, $templateRequest, $sce, $compile) {

        var helpDataSource = $sce.getTrustedResourceUrl('/assembly/' + ASPOSE_PRODUCTNAME + 'HelpDataSource.html');
        $templateRequest(helpDataSource).then(function (template) {
            $compile($("#help-dialog-datasource > .modal-dialog > .modal-content").html(template).contents())($scope);
        });

        var helpTemplate = $sce.getTrustedResourceUrl('/assembly/' + ASPOSE_PRODUCTNAME + 'HelpTemplate.html');
        $templateRequest(helpTemplate).then(function (template) {
            $compile($("#help-dialog-template > .modal-dialog > .modal-content").html(template).contents())($scope);
        });


        $scope.uploadTemplateFile = function () {
            $scope.uploadFile(
                $scope.templateFile.file,
                function (response) {
                    setTimeout(function () {
                        $scope.templateFile.progress = 100;
                        $scope.stage = 2;
                        $scope.showed = Math.max($scope.showed, 2);
                        $scope.$apply();
                    }, 1500);
                },
                function (response) {
                    alertError(response);
                },
                function (e) {
                    $scope.templateFile.loadedSize = e.loaded;
                    $scope.templateFile.totalSize = e.total;
                    $scope.templateFile.progress = Math.min(99, parseInt(100.0 * e.loaded / e.total));
                }
            );
        };

        $scope.uploadDatasourceFile = function () {
            $scope.uploadFile(
                $scope.datasourceFile.file,
                function () {
                    setTimeout(function () {
                        $scope.datasourceFile.progress = 100;
                        $scope.stage = 3;
                        $scope.showed = Math.max($scope.showed, 3);
                        $scope.$apply();
                    }, 1500);
                },
                function (response) {
                    alertError(response);
                },
                function (e) {
                    $scope.datasourceFile.loadedSize = e.loaded;
                    $scope.datasourceFile.totalSize = e.total;
                    $scope.datasourceFile.progress = Math.min(99, parseInt(100.0 * e.loaded / e.total));
                }
            );
        };

        $scope.assembleDocument = function () {

            var url = ASPOSE_ASSEMBLY_API + 'Assemble?' + $.param({
                productName: ASPOSE_PRODUCTNAME,
                folderName: $scope.folderName,
                templateFilename: $scope.templateFile.file.name,
                datasourceFilename: $scope.datasourceFile.file.name,
                datasourceName: $scope.datasourceName,
                datasourceTableIndex: $scope.datasourceTableIndex,
                delimiter: encodeURI($scope.delimiter)
            });

            $scope.isLoading = true;
            $http({
                method: 'POST',
                url: url,
                responseType: "application/json"
            }).then(
                function (response) {
                    $scope.isLoading = false;
                    if (response.data.StatusCode === 200) {
                        $('#AssemblyMessage').hide();
                        var downloadLink = ASPOSE_FILEDOWNLOADLINK +
                            $.param({
                                FolderName: response.data.FolderName,
                                FileName: response.data.FileName
                            });
                        window.location.href = downloadLink;
                    } else
                        alertError(response);
                },
                function (response) {
                    $scope.isLoading = false;
                    alertError(response);
                }
            );
        };

        $scope.uploadFile = function (file, success, error, progress) {
            Upload.upload({
                method: 'POST',
                url: ASPOSE_ASSEMBLY_API + 'Upload?' + $.param({
                    folderName: $scope.folderName
                }),
                data: {
                    file: file
                }
            }).then(success, error, progress);
        };

        $scope.showStage = function (stagenumber) {
            $('#AssemblyMessage').hide();
            $scope.stage = stagenumber;
        };

        $scope.setDelimiter = function (delimiter, buttontext = null) {
            $scope.delimiter = delimiter;
            if (buttontext === null)
                buttontext = delimiter;
            $('#delimiter')[0].innerText = buttontext;
        };

        $scope.$watch('datasourceFile.file',
            function (newval, oldval) {
                if (newval !== undefined) {
                    var re = /(?:\.([^.]+))?$/;
                    var extension = re.exec(newval.name)[1].toLowerCase();
                    switch (extension) {
                        case "xml":
                        case "json":
                            $scope.showTableIndex = false;
                            $scope.showDelimiter = false;
                            break;
                        case "csv":
                            $scope.showTableIndex = false;
                            $scope.showDelimiter = true;
                            break;
                        default:
                            $scope.showTableIndex = true;
                            $scope.showDelimiter = false;
                            break;
                    }

                    if (ASPOSE_PRODUCTNAME == "cells") {
                        $scope.showTableIndex = false;
                        $scope.showDelimiter = false;
                    }
                }
            });

        $scope.start = function () {
            $scope.folderName = randomString();
            $scope.templateFile = {};
            $scope.templateError = null;
            $scope.datasourceFile = {};
            $scope.datasourceError = null;
            $scope.datasourceTableIndex = 0;
            $scope.datasourceName = "";
            $scope.stage = 1;
            $scope.isLoading = false;
            $scope.showed = 1;
            $scope.delimiter = ",";
            $scope.showDelimiter = false;
            $scope.showTableIndex = false;
            $('#AssemblyMessage').hide();
        };

        $scope.start();
    }

    function randomString() {
        return Math.random().toString(36).substring(2)
            + Math.random().toString(36).substring(2)
            + Math.random().toString(36).substring(2);
    }

    function alertError(response) {
        var m = $('#AssemblyMessage');
        m.removeClass('alert-success');
        m.addClass('alert-danger');
        if (response.data === null)
            m.text(response.xhrStatus);
        else
            m.text(response.data.Status);
        m.show();
    }

    function alertSuccess(message) {
        var m = $('#AssemblyMessage');
        m.removeClass('alert-danger');
        m.addClass('alert-success');
        m.text(message);
        m.show();
    }

    angular.module('AsposeAssemblyApp').controller('Main', main);
})();
