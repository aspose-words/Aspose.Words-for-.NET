(function ($) {
    $.metadata = function (data, fileId, fileName) {
        const BUILT_IN = 1;
        const CUSTOM   = 2;
        const ITEMS_ON_PAGE = 10;
        const TYPES = {
            0: 'Boolean',
            1: 'DateTime',
            2: 'Double',
            3: 'Number',
            4: 'String',
            5: 'StringArray',
            6: 'ObjectArray',
            7: 'ByteArray',
            8: 'Other',
        };

        var mode = BUILT_IN;
        var page = 1;

        var itemBackup = {};

        var checkType = function(value, type) {
            const TYPE_ERROR = 'Value and Type do not match, please try again!';
            var valueTmp;

            switch (type) {
                case 4: // String
                case 5: // StringArray
                case 6: // ObjectArray
                    break;
                case 1: // DateTime
                    if (isNaN(Date.parse(value))) {
                        window.alert(TYPE_ERROR);
                        return null;
                    }
                    break;
                case 0: // Boolean
                    valueTmp = value.toLowerCase();
                    if (valueTmp === 'true') {
                        value = true;
                    } else if (valueTmp === 'false') {
                        value = false;
                    } else {
                        window.alert(TYPE_ERROR);
                        return null;
                    }
                    break;
                case 2: // Double
                case 3: // Number
                    valueTmp = parseFloat(value);
                    if (isNaN(valueTmp) || !value.match(/^[+-]?\d+(\.\d+)?$/)) {
                        window.alert(TYPE_ERROR);
                        return null;
                    }
                    value = type === 2 ? valueTmp : Math.floor(valueTmp);
                    break;
                case 7: // ByteArray
                    valueTmp = value.split(',');
                    if (valueTmp.reduce(function(curr, item) {
                        return (curr === true && item.trim().match(/^[+-]?\d+(\.\d+)?$/));
                    }, true) === false) {
                        window.alert(TYPE_ERROR);
                        return null;
                    }
                    break;
                default:
                    window.alert('Unknown type was specified, please try again!');
                    return null;
            }
            return value;
        };

        var reset = function() {
            $('.filedrop').removeClass('hidden');
            $('#metadata').addClass('hidden');
            fileDrop.reset();
        };

        var cancel = function() {
            if (itemBackup.id) {
                var tr = $('tr#item-' + itemBackup.id);
                if (tr.length > 0) {
                    var tdValue  = tr.find('td[data-item-value]');
                    var tdButton = tr.find('td[data-item-button]');
                    if (tdValue.length && tdButton.length) {
                        itemBackup.button.on('click', onClickModify);
                        tdValue.empty().append(itemBackup.value);
                        tdButton.empty().append(itemBackup.button);
                    }
                }
            }
        };

        var onClickModify = function(evt) {
            evt.preventDefault();
            evt.stopPropagation();
            var id = parseInt(event.target.getAttribute('data-target'));
            modify(id);
        };

        var onClickRemove = function(evt) {
            evt.preventDefault();
            evt.stopPropagation();
            var id = parseInt(event.target.getAttribute('data-target'));
            remove(id);
        };

        var onClickUpdate = function(evt) {
            if (evt) {
                evt.preventDefault();
                evt.stopPropagation();
            }
            var value = checkType($('input#item-value-' + itemBackup.id).val(), itemBackup.type);
            if (value === null) {
                return;
            }
            var found = false;
            var i;
            for (i = 0; i < data.BuiltIn.length; i++) {
                if (data.BuiltIn[i].Id === itemBackup.id) {
                    data.BuiltIn[i].Value = value;
                    found = true;
                    break;
                }
            }
            if (!found) {
                for (i = 0; i < data.Custom.length; i++) {
                    if (data.Custom[i].Id === itemBackup.id) {
                        data.Custom[i].Value = value;
                        found = true;
                        break;
                    }
                }
            }
            if (found) {
                itemBackup.value = $('<span>' + value + '</span>');
            }
            cancel();
        };

        var onClickCancel = function(evt) {
            evt.preventDefault();
            evt.stopPropagation();
            cancel();
        };

        var modify = function(id) {
            // cancel previous editing
            cancel();

            var tr = $('tr#item-' + id);
            if (tr.length > 0) {
                var tdValue  = tr.find('td[data-item-value]');
                var tdButton = tr.find('td[data-item-button]');
                var tdType   = tr.find('td[data-item-type]');
                if (tdValue.length && tdButton.length && tdType.length) {
                    var value = tdValue.find('span').html();

                    // save the backup values
                    itemBackup.id     = id;
                    itemBackup.value  = tdValue.children().detach();
                    itemBackup.button = tdButton.children().detach();
                    itemBackup.type   = parseInt(tdType.attr('data-item-type'));
                    itemBackup.button.off('click');

                    var input = $('<input id="item-value-' + id + '" type="text" value="" />');
                    input.val(value.toString());
                    input.on('keydown', function(evt) {
                        if (evt.keyCode === 13) {
                            evt.preventDefault();
                            evt.stopPropagation();
                            onClickUpdate();
                        }
                    });
                    tdValue.append(input);

                    var btnUpdate = $('<input type="image" src="../img/grid-update-ico.png" title="Update" alt="Update" style="border-width:0" />');
                    btnUpdate.on('click', onClickUpdate);
                    tdButton.append(btnUpdate);
                    tdButton.append('&nbsp;');

                    var btnCancel = $('<input type="image" src="../img/grid-cancel-ico.png" title="Cancel" alt="Cancel" style="border-width:0" />');
                    btnCancel.on('click', onClickCancel);
                    tdButton.append(btnCancel);
                }
            }
        };

        var remove = function(id) {
            // cancel previous editing
            cancel();

            var tr = $('tr#item-' + id);
            if (tr.length > 0) {
                var found = false;
                var i;
                for (i = 0; i < data.Custom.length; i++) {
                    if (data.Custom[i].Id === id) {
                        found = true;
                        break;
                    }
                }
                if (found) {
                    data.Custom.splice(i, 1);
                    display();
                }
            }
        };

        var create = function() {
            var name = $('#new-property-name').val();
            if (!name || name.match(/^\s+$/)) {
                window.alert('Property name can not be empty, please try again!');
                return;
            }
            var type = parseInt($('#new-property-type').val());
            var value = checkType($('#new-property-value').val(), type);
            if (value !== null) {
                var id = data.BuiltIn.length;
                var index;
                do {
                    id += 1;
                    index = data.Custom.findIndex(prop => prop.Id === id);
                } while (index !== -1);
                data.Custom.push({
                    Id              : id,
                    Name            : name,
                    Value           : value,
                    Type            : type,
                    LinkSource      : '',
                    IsLinkToContent : false
                });

                $('#new-property-name').val('');
                $('#new-property-value').val('');
                $('#new-property-type').val('4');

                page = Math.ceil(data.Custom.length / ITEMS_ON_PAGE); // select the last page
                display();
            }
        }

        var display = function() {
            var begin = (page - 1) * ITEMS_ON_PAGE;
            var end   = begin + ITEMS_ON_PAGE;
            var items = mode === BUILT_IN ? data.BuiltIn : data.Custom;
            var pages = Math.ceil(items.length / ITEMS_ON_PAGE);
            if (items.length < end) {
                end = items.length;
            }
            items = items.slice(begin, end);

            var tbody = $('<tbody>');
            var tr, td;
            for (var i in items) {
                td = $('<td data-item-button>');
                var btn = $('<input type="image" src="../img/grid-edit-ico.png" alt="Edit" title="Edit" style="border-width:0" data-target="' + items[i].Id + '" />');
                btn.on('click', onClickModify);
                td.append(btn);
                if (mode === CUSTOM) {
                    btn = $('<input type="image" src="../img/grid-delete-ico.png" alt="Delete" title="Delete" style="border-width:0" data-target="' + items[i].Id + '" />');
                    btn.on('click', onClickRemove);
                    td.append('&nbsp;');
                    td.append(btn);
                }

                tr = $('<tr id="item-' + items[i].Id + '">');
                tr.append($('<td><span>' + items[i].Name + '</span></td>'));
                tr.append($('<td data-item-value><span>' + items[i].Value + '</span></td>'));
                tr.append($('<td data-item-type="' + items[i].Type + '"><span>' + TYPES[items[i].Type] + '</span></td>'));
                tr.append(td);
                tbody.append(tr);
            }
            // list of pages
            if (pages > 1) {
                tr = $('<tr>');
                td = $('<td colspan="4">');
                var table = $('<table border="0">');
                td.append(table);
                tr.append(td);
                tbody.append(tr);

                tr = $('<tr>');
                table.append(tr);

                var pageLink;
                for (i = 1; i <= pages; i++) {
                    td = $('<td>');
                    if (i === page) { // current page
                        pageLink = $('<span>' + i + '</span>');
                    } else { // other page
                        pageLink = $('<a href="javascript:void(0)">' + i + '</a>');
                        pageLink.on('click', function(el) {
                            page = parseInt(el.target.innerHTML, 10);
                            display();
                        });
                    }
                    td.append(pageLink);
                    tr.append(td);
                }
            };

            $('table#metadata-props > tbody').remove();
            $('table#metadata-props').append(tbody);
        };

        var onSave = function() {
            showLoader();
			const url = o.UIBasePath + 'api/AsposeWordsMetadata/download';

            const properties = {
                BuiltIn: data.BuiltIn.map(p => {
                    return {
                        Name            : p.Name,
                        Value           : p.Value,
                        Type            : p.Type,
                        LinkSource      : p.LinkSource,
                        IsLinkToContent : p.IsLinkToContent
                    };
                }),
                Custom: data.Custom.map(p => {
                    return {
                        Name            : p.Name,
                        Value           : p.Value,
                        Type            : p.Type,
                        LinkSource      : p.LinkSource,
                        IsLinkToContent : p.IsLinkToContent
                    };
                })
            };

            $.ajax({
                method: 'POST',
                url: url,
                data: JSON.stringify({
                    id: fileId,
                    FileName: fileName,
                    properties
                }),
                contentType: 'application/json',
                cache: false,
                timeout: 600000,
                success: workSuccess,
                error: (err) => {
                    if (err.data !== undefined && err.data.Status !== undefined)
                        showAlert(err.data.Status);
                    else
                        showAlert("Error " + err.status + ": " + err.statusText);
                }
            });
        };

        var onClearAll = function() {
            if (window.confirm('Are you sure you want to clear all metadata?')) {
                showLoader();
                const url = o.UIBasePath + 'api/AsposeWordsMetadata/clear';

                $.ajax({
                    method: 'POST',
                    url: url,
                    data: JSON.stringify({
                        id: fileId,
                        FileName: fileName
                    }),
                    contentType: 'application/json',
                    cache: false,
                    timeout: 600000,
                    success: workSuccess,
                    error: (err) => {
                        if (err.data !== undefined && err.data.Status !== undefined)
                            showAlert(err.data.Status);
                        else
                            showAlert("Error " + err.status + ": " + err.statusText);
                    }
                });
            }
        };

        $('#metadata-mode-builtin').on('click', function() {
            if (mode !== BUILT_IN) {
                mode = BUILT_IN;
                display();
                $('#new-property').addClass('hidden');
            }
        });
        $('#metadata-mode-custom').on('click', function() {
            if (mode !== CUSTOM) {
                mode = CUSTOM;
                display();
                $('#new-property').removeClass('hidden');
            }
        });

        // set Ids for input data
        var id = 1;
        var i;
        for (i = 0; i < data.BuiltIn.length; i++) {
            data.BuiltIn[i].Id = id++;
        }
        for (i = 0; i < data.Custom.length; i++) {
            data.Custom[i].Id = id++;
        }

        $('#btn-save').on('click', onSave);
        $('#btn-clear-all').on('click', onClearAll);
        $('#btn-cancel').on('click', reset);
        $('#new-property-name').on('keydown', function(evt) {
            if (evt.keyCode === 13) {
                evt.preventDefault();
                evt.stopPropagation();
                create();
            }
        });
        $('#new-property-value').on('keydown', function(evt) {
            if (evt.keyCode === 13) {
                evt.preventDefault();
                evt.stopPropagation();
                create();
            }
        });
        $('#new-property-save').on('click', create);

        $('.filedrop').addClass('hidden');
        $('#metadata-edit').removeClass('hidden');
        hideLoader();
        display();
    };
})(jQuery);
