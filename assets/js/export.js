/* global mdpExportConfig */
(function($) {
    'use strict';

    if (typeof $ === 'undefined') {
        return;
    }

    var config = window.mdpExportConfig || {};
    var i18n = config.i18n || {};
    console.log('mdp-export: loaded');
    var mdpSaveTimer = null;
    var mdpRulesSaveTimer = null;
    var mdpFormulasSaveTimer = null;
    var mdpLastFormulaInput = null;

    function setFormId(value) {
        $('#mdp_export_form_id_csv').val(value);
        $('#mdp_export_form_id_save').val(value);
        config.formId = value;
    }

    function renderFormulaTags(keys) {
        var $list = $('#mdp-formula-tags-list');
        $list.empty();
        if (!keys || !keys.length) {
            return;
        }
        keys.forEach(function(key) {
            var tag = $('<button type="button" class="mdp-formula-tag"></button>').text(key).attr('data-key', key);
            $list.append(tag);
        });
    }

    function getDateFormatOptions() {
        return config.dateFormats || [
            { value: '', label: 'Original' },
            { value: 'd.m.Y', label: 'TT.MM.JJJJ' }
        ];
    }

    function buildDateFormatSelect(current) {
        var dateSelect = $('<select class="mdp-field-date-format"></select>');
        getDateFormatOptions().forEach(function(option) {
            var opt = $('<option></option>').val(option.value).text(option.label);
            if (option.value === current) {
                opt.prop('selected', true);
            }
            dateSelect.append(opt);
        });
        return dateSelect;
    }

    function renderFields(fields, dateCandidates) {
        var $list = $('#mdp-export-fields-list');
        $list.empty();
        if (!fields || !fields.length) {
            $list.append('<tr class="mdp-export-empty-row"><td colspan="5">' + (i18n.no_fields || 'Keine Felder gefunden.') + '</td></tr>');
            renderFormulaTags([]);
            return;
        }
        var dateCandidateSet = Array.isArray(dateCandidates) ? dateCandidates : [];
        var tagKeys = [];
        fields.forEach(function(field) {
            var key = field.key || '';
            var label = field.label || key;
            var include = field.include !== false;
            var dateFormat = field.date_format || '';
            var isCustom = field.custom === true;
            var isFormula = field.formula === true;
            var isDateCandidate = key === 'created_at' || dateCandidateSet.indexOf(key) !== -1;
            var showDateFormat = isDateCandidate || dateFormat;
            if (key && !isFormula) {
                tagKeys.push(key);
            }
            var item = $('<tr class="mdp-export-field-item" />').attr('data-key', key);
            if (isCustom) {
                item.attr('data-custom', '1');
            }
            if (isFormula) {
                item.attr('data-formula', '1');
            }
            item.append('<td class="mdp-export-col-handle"><span class="mdp-field-handle" aria-hidden="true">≡</span></td>');
            var keyCell = $('<td class="mdp-field-key"></td>').text(key);
            if (isCustom) {
                keyCell.addClass('mdp-export-custom-label');
                keyCell.append('<button type="button" class="mdp-export-delete-field" aria-label="' + (i18n.delete_field || 'Feld löschen') + '"><span class="dashicons dashicons-no-alt" aria-hidden="true"></span></button>');
            }
            if (isFormula) {
                keyCell.addClass('mdp-export-formula-label');
            }
            item.append(keyCell);
            var labelCell = $('<td class="mdp-field-label-cell"></td>');
            var labelInput = $('<input type="text" class="mdp-field-label" />').val(label);
            labelCell.append(labelInput);
            item.append(labelCell);
            var dateCell = $('<td class="mdp-field-dateformat-cell"></td>');
            if (showDateFormat) {
                var dateSelect = buildDateFormatSelect(dateFormat);
                dateCell.append(dateSelect);
            }
            item.append(dateCell);
            var includeCell = $('<td class="mdp-export-col-include"></td>');
            var includeInput = $('<input type="checkbox" class="mdp-field-include" />').prop('checked', include);
            includeCell.append(includeInput);
            item.append(includeCell);
            $list.append(item);
        });
        renderFormulaTags(tagKeys);
        if ($.fn.sortable) {
            $list.sortable({
                helper: function(e, ui) {
                    ui.children().each(function() {
                        $(this).width($(this).width());
                    });
                    return ui;
                },
                handle: '.mdp-field-handle, .mdp-field-key',
                update: function() {
                    updatePayload(true);
                }
            });
        }
        updatePayload(false);
    }

    function updatePayload(shouldSave) {
        var payload = [];
        $('#mdp-export-fields-list .mdp-export-field-item').each(function() {
            var $item = $(this);
            var dateValue = '';
            var $dateSelect = $item.find('.mdp-field-date-format');
            if ($dateSelect.length) {
                dateValue = $dateSelect.val() || '';
            }
            payload.push({
                key: $item.data('key'),
                label: $item.find('.mdp-field-label').val(),
                date_format: dateValue,
                include: $item.find('.mdp-field-include').is(':checked'),
                custom: $item.data('custom') === 1 || $item.data('custom') === '1',
                formula: $item.data('formula') === 1 || $item.data('formula') === '1'
            });
        });
        $('#mdp_export_fields_payload').val(JSON.stringify(payload));
        if (shouldSave) {
            scheduleSave();
        }
    }

    function scheduleSave() {
        if (mdpSaveTimer) {
            clearTimeout(mdpSaveTimer);
        }
        mdpSaveTimer = setTimeout(saveFieldsAjax, 700);
    }

    function showSaveToast() {
        var $toast = $('#mdp-export-toast');
        if (!$toast.length) {
            $toast = $('<div id="mdp-export-toast" class="mdp-export-toast" role="status" aria-live="polite">' + (i18n.saved || 'Gespeichert') + '</div>');
            $('body').append($toast);
        }
        $toast.addClass('is-visible');
        clearTimeout($toast.data('timeout'));
        var timeoutId = setTimeout(function() {
            $toast.removeClass('is-visible');
        }, 2000);
        $toast.data('timeout', timeoutId);
    }

    function saveFieldsAjax() {
        var payload = $('#mdp_export_fields_payload').val();
        var formId = $('#mdp_export_form_id_save').val();
        if (!formId) {
            return;
        }
        $.post(config.ajaxUrl, {
            action: 'mdp_save_export_fields_ajax',
            nonce: config.nonce,
            form_id: formId,
            fields_payload: payload
        }).done(function(response) {
            if (response && response.success) {
                showSaveToast();
                loadPreview(formId);
            }
        });
    }

    function loadFields(formId) {
        $.post(config.ajaxUrl, {
            action: 'mdp_get_export_fields',
            nonce: config.nonce,
            form_id: formId
        }).done(function(response) {
            if (response && response.success && response.data && response.data.fields) {
                renderFields(response.data.fields, response.data.date_candidates || []);
            } else {
                renderFields([], []);
            }
        }).fail(function() {
            renderFields([], []);
        });
    }

    function renderPreview(headers, rows) {
        var $table = $('.mdp-export-preview-table');
        var $thead = $table.find('thead');
        var $tbody = $table.find('tbody');
        $thead.empty();
        $tbody.empty();

        if (!headers || !headers.length) {
            $thead.append('<tr><th>' + (i18n.no_columns || 'Keine Spalten ausgewählt.') + '</th></tr>');
            $tbody.append('<tr><td>' + (i18n.no_columns_body || 'Bitte mindestens eine Export-Spalte aktivieren.') + '</td></tr>');
            return;
        }

        var headerRow = $('<tr></tr>');
        headers.forEach(function(label) {
            headerRow.append($('<th></th>').text(label));
        });
        $thead.append(headerRow);

        if (!rows || !rows.length) {
            $tbody.append('<tr><td colspan="' + headers.length + '">' + (i18n.no_entries || 'Keine Einträge gefunden.') + '</td></tr>');
            return;
        }

        rows.forEach(function(row) {
            var $row = $('<tr></tr>');
            row.forEach(function(cell) {
                $row.append($('<td></td>').text(cell));
            });
            $tbody.append($row);
        });
    }

    function loadPreview(formId) {
        $.post(config.ajaxUrl, {
            action: 'mdp_get_export_preview',
            nonce: config.nonce,
            form_id: formId
        }).done(function(response) {
            if (response && response.success) {
                renderPreview(response.data.headers || [], response.data.rows || []);
            }
        });
    }

    function renderRules(rules) {
        var $body = $('#mdp-export-rules-body');
        $body.empty();
        if (!rules || !rules.length) {
            return;
        }
        rules.forEach(function(rule) {
            var row = $('<tr class="mdp-export-rule-row"></tr>');
            var findInput = $('<input type="text" class="mdp-export-rule-find">').val(rule.find || '');
            var replaceInput = $('<input type="text" class="mdp-export-rule-replace">').val(rule.replace || '');
            row.append($('<td></td>').append(findInput));
            row.append($('<td></td>').append(replaceInput));
            row.append('<td><button type="button" class="button-link mdp-export-rule-remove" aria-label="' + (i18n.delete_rule || 'Regel löschen') + '"><span class="dashicons dashicons-no-alt" aria-hidden="true"></span></button></td>');
            $body.append(row);
        });
    }

    function updateRulesPayload(shouldSave) {
        var payload = [];
        $('#mdp-export-rules-body .mdp-export-rule-row').each(function() {
            var $row = $(this);
            var find = $row.find('.mdp-export-rule-find').val() || '';
            var replace = $row.find('.mdp-export-rule-replace').val() || '';
            if (find.trim() === '') {
                return;
            }
            payload.push({
                find: find,
                replace: replace
            });
        });
        if (shouldSave) {
            scheduleRulesSave(payload);
        }
    }

    function scheduleRulesSave(payload) {
        if (mdpRulesSaveTimer) {
            clearTimeout(mdpRulesSaveTimer);
        }
        mdpRulesSaveTimer = setTimeout(function() {
            saveRulesAjax(payload);
        }, 700);
    }

    function saveRulesAjax(payload) {
        var formId = $('#mdp_export_form_id_save').val();
        if (!formId) {
            return;
        }
        $.post(config.ajaxUrl, {
            action: 'mdp_save_export_rules_ajax',
            nonce: config.nonce,
            form_id: formId,
            rules_payload: JSON.stringify(payload || [])
        }).done(function(response) {
            if (response && response.success) {
                showSaveToast();
                loadPreview(formId);
            }
        });
    }

    function loadRules(formId) {
        $.post(config.ajaxUrl, {
            action: 'mdp_get_export_rules',
            nonce: config.nonce,
            form_id: formId
        }).done(function(response) {
            if (response && response.success) {
                renderRules(response.data.rules || []);
            } else {
                renderRules([]);
            }
        }).fail(function() {
            renderRules([]);
        });
    }

    function renderFormulas(formulas) {
        var $body = $('#mdp-export-formulas-body');
        $body.empty();
        if (!formulas || !formulas.length) {
            return;
        }
        formulas.forEach(function(item) {
            var row = $('<tr class="mdp-export-formula-row"></tr>').attr('data-key', item.key || '');
            var labelInput = $('<input type="text" class="mdp-export-formula-label">').val(item.label || '');
            var formulaInput = $('<input type="text" class="mdp-export-formula-expression">').val(item.formula || '');
            row.append($('<td></td>').append(labelInput));
            row.append($('<td></td>').append(formulaInput));
            row.append('<td><button type="button" class="button-link mdp-export-formula-remove" aria-label="' + (i18n.delete_formula || 'Formel löschen') + '"><span class="dashicons dashicons-no-alt" aria-hidden="true"></span></button></td>');
            $body.append(row);
        });
    }

    function updateFormulasPayload(shouldSave) {
        var payload = [];
        $('#mdp-export-formulas-body .mdp-export-formula-row').each(function() {
            var $row = $(this);
            var key = $row.data('key') || '';
            var label = $row.find('.mdp-export-formula-label').val() || '';
            var formula = $row.find('.mdp-export-formula-expression').val() || '';
            if (!key || formula.trim() === '') {
                return;
            }
            payload.push({
                key: key,
                label: label,
                formula: formula
            });
        });
        if (shouldSave) {
            scheduleFormulasSave(payload);
        }
    }

    function scheduleFormulasSave(payload) {
        if (mdpFormulasSaveTimer) {
            clearTimeout(mdpFormulasSaveTimer);
        }
        mdpFormulasSaveTimer = setTimeout(function() {
            saveFormulasAjax(payload);
        }, 700);
    }

    function saveFormulasAjax(payload) {
        var formId = $('#mdp_export_form_id_save').val();
        if (!formId) {
            return;
        }
        $.post(config.ajaxUrl, {
            action: 'mdp_save_export_formulas_ajax',
            nonce: config.nonce,
            form_id: formId,
            formulas_payload: JSON.stringify(payload || [])
        }).done(function(response) {
            if (response && response.success) {
                showSaveToast();
                loadFields(formId);
                loadPreview(formId);
            }
        });
    }

    function loadFormulas(formId) {
        $.post(config.ajaxUrl, {
            action: 'mdp_get_export_formulas',
            nonce: config.nonce,
            form_id: formId
        }).done(function(response) {
            if (response && response.success) {
                renderFormulas(response.data.formulas || []);
            } else {
                renderFormulas([]);
            }
        }).fail(function() {
            renderFormulas([]);
        });
    }

    $(document).on('change', '#mdp_export_form_id', function() {
        var formId = $(this).val();
        setFormId(formId);
        loadFields(formId);
        loadPreview(formId);
        loadRules(formId);
        loadFormulas(formId);
        $.post(config.ajaxUrl, {
            action: 'mdp_save_export_last_form',
            nonce: config.nonce,
            form_id: formId
        });
    });

    $(document).on('change', '#mdp_export_format', function() {
        var format = $(this).val();
        $.post(config.ajaxUrl, {
            action: 'mdp_save_export_last_format',
            nonce: config.nonce,
            format: format
        });
    });

    $(document).on('click', '#mdp_export_add_field', function() {
        var rawName = $('#mdp_export_new_field').val();
        var name = rawName ? rawName.trim() : '';
        if (!name) {
            return;
        }
        var key = name.toLowerCase()
            .replace(/\s+/g, '_')
            .replace(/[^a-z0-9_]/g, '')
            .replace(/^_+|_+$/g, '');
        if (!key) {
            key = 'custom_' + Date.now();
        }
        if ($('#mdp-export-fields-list .mdp-export-field-item[data-key="' + key + '"]').length) {
            $('#mdp_export_new_field').val('');
            return;
        }
        var item = $('<tr class="mdp-export-field-item" />').attr('data-key', key).attr('data-custom', '1');
        item.append('<td class="mdp-export-col-handle"><span class="mdp-field-handle" aria-hidden="true">≡</span></td>');
        var keyCell = $('<td class="mdp-field-key mdp-export-custom-label"></td>').text(key);
        keyCell.append('<button type="button" class="mdp-export-delete-field" aria-label="' + (i18n.delete_field || 'Feld löschen') + '"><span class="dashicons dashicons-no-alt" aria-hidden="true"></span></button>');
        item.append(keyCell);
        var labelCell = $('<td class="mdp-field-label-cell"></td>');
        var labelInput = $('<input type="text" class="mdp-field-label" />').val(name);
        labelCell.append(labelInput);
        item.append(labelCell);
        var dateCell = $('<td class="mdp-field-dateformat-cell"></td>');
        item.append(dateCell);
        var includeCell = $('<td class="mdp-export-col-include"></td>');
        var includeInput = $('<input type="checkbox" class="mdp-field-include" />').prop('checked', true);
        includeCell.append(includeInput);
        item.append(includeCell);
        $('#mdp-export-fields-list').append(item);
        $('#mdp_export_new_field').val('');
        updatePayload(true);
    });

    $(document).on('click', '.mdp-export-delete-field', function(e) {
        e.preventDefault();
        $(this).closest('.mdp-export-field-item').remove();
        updatePayload(true);
    });

    $(document).on('click', '#mdp_export_add_rule', function() {
        var $body = $('#mdp-export-rules-body');
        var row = $('<tr class="mdp-export-rule-row"></tr>');
        var findInput = $('<input type="text" class="mdp-export-rule-find">').val('');
        var replaceInput = $('<input type="text" class="mdp-export-rule-replace">').val('');
        row.append($('<td></td>').append(findInput));
        row.append($('<td></td>').append(replaceInput));
        row.append('<td><button type="button" class="button-link mdp-export-rule-remove" aria-label="' + (i18n.delete_rule || 'Regel löschen') + '"><span class="dashicons dashicons-no-alt" aria-hidden="true"></span></button></td>');
        $body.append(row);
        updateRulesPayload(true);
    });

    $(document).on('input', '.mdp-export-rule-find, .mdp-export-rule-replace', function() {
        updateRulesPayload(true);
    });

    $(document).on('click', '.mdp-export-rule-remove', function(e) {
        e.preventDefault();
        $(this).closest('.mdp-export-rule-row').remove();
        updateRulesPayload(true);
    });

    $(document).on('click', '#mdp_export_add_formula', function() {
        var label = $('#mdp_export_new_formula_label').val() || '';
        label = label.trim();
        if (!label) {
            return;
        }
        var key = label.toLowerCase()
            .replace(/\s+/g, '_')
            .replace(/[^a-z0-9_]/g, '')
            .replace(/^_+|_+$/g, '');
        if (!key) {
            key = 'formula_' + Date.now();
        }
        if ($('#mdp-export-formulas-body .mdp-export-formula-row[data-key="' + key + '"]').length) {
            return;
        }
        var row = $('<tr class="mdp-export-formula-row"></tr>').attr('data-key', key);
        var labelInput = $('<input type="text" class="mdp-export-formula-label">').val(label);
        var formulaInput = $('<input type="text" class="mdp-export-formula-expression">').val('');
        row.append($('<td></td>').append(labelInput));
        row.append($('<td></td>').append(formulaInput));
        row.append('<td><button type="button" class="button-link mdp-export-formula-remove" aria-label="' + (i18n.delete_formula || 'Formel löschen') + '"><span class="dashicons dashicons-no-alt" aria-hidden="true"></span></button></td>');
        $('#mdp-export-formulas-body').append(row);
        $('#mdp_export_new_formula_label').val('');
        formulaInput.focus();
        mdpLastFormulaInput = formulaInput[0];
        updateFormulasPayload(true);
    });

    $(document).on('input', '.mdp-export-formula-label, .mdp-export-formula-expression', function() {
        updateFormulasPayload(true);
    });

    $(document).on('click', '.mdp-export-formula-remove', function(e) {
        e.preventDefault();
        $(this).closest('.mdp-export-formula-row').remove();
        updateFormulasPayload(true);
    });

    $(document).on('focus', '.mdp-export-formula-expression', function() {
        mdpLastFormulaInput = this;
    });

    $(document).on('click', '.mdp-formula-tag', function() {
        var key = $(this).data('key');
        if (!key) {
            return;
        }
        var input = mdpLastFormulaInput || document.querySelector('.mdp-export-formula-expression');
        if (!input) {
            return;
        }
        var value = input.value || '';
        var insert = key;
        if (value.trim() !== '') {
            insert = ' & ' + key;
        }
        var start = input.selectionStart || value.length;
        var end = input.selectionEnd || value.length;
        var next = value.slice(0, start) + insert + value.slice(end);
        input.value = next;
        var caret = start + insert.length;
        input.setSelectionRange(caret, caret);
        input.focus();
        if ($(input).hasClass('mdp-export-formula-expression')) {
            updateFormulasPayload(true);
        }
    });

    $(document).on('click', '.mdp-export-tab', function() {
        var $tab = $(this);
        var target = $tab.data('tab');
        $('.mdp-export-tab').removeClass('is-active');
        $tab.addClass('is-active');
        $('.mdp-export-tab-panel').removeClass('is-active');
        $('#' + target).addClass('is-active');
    });

    $(document).on('input', '.mdp-field-label', function() {
        updatePayload(true);
    });
    $(document).on('change', '.mdp-field-date-format', function() {
        updatePayload(true);
    });
    $(document).on('change', '.mdp-field-include', function() {
        updatePayload(true);
    });
    $('#mdp-export-fields-form').on('submit', function() {
        updatePayload(true);
    });

    if (!config.formId) {
        config.formId = $('#mdp_export_form_id_save').val() || $('#mdp_export_form_id').val() || '';
    }
    setFormId(config.formId);
    if (config.formId) {
        loadFields(config.formId);
        loadPreview(config.formId);
        loadRules(config.formId);
        loadFormulas(config.formId);
    }
})(jQuery);
