jQuery(function($) {
    var config = window.mdpConversionsConfig || {};
    var $formSelect = $('#mdp_conversion_form_id');
    var $consentSelect = $('#mdp_consent_tool');

    function updateConsentFields() {
        var tool = $consentSelect.val();
        $('.mdp-consent-fields').hide();
        if (tool && tool !== 'none') {
            $('.mdp-consent-fields[data-tool=\"' + tool + '\"]').show();
        }
    }

    function showToast() {
        var $toast = $('#mdp-export-toast');
        if (!$toast.length) {
            $toast = $('<div id="mdp-export-toast" class="mdp-export-toast" role="status" aria-live="polite">' + (config.i18n && config.i18n.saved ? config.i18n.saved : 'Gespeichert') + '</div>');
            $('body').append($toast);
        }
        $toast.addClass('is-visible');
        clearTimeout($toast.data('timeout'));
        var timeoutId = setTimeout(function() {
            $toast.removeClass('is-visible');
        }, 2000);
        $toast.data('timeout', timeoutId);
    }

    function reloadWithForm(formId) {
        var url = new URL(window.location.href);
        url.searchParams.set('form_id', formId);
        window.location.href = url.toString();
    }

    $formSelect.on('change', function() {
        var formId = $(this).val();
        var $form = $(this).closest('form');
        if (config.ajaxUrl && config.nonce) {
            $.post(config.ajaxUrl, {
                action: 'mdp_save_conversion_last_form',
                nonce: config.nonce,
                form_id: formId
            });
        }
        if (!$form.length) {
            reloadWithForm(formId);
        }
    });

    $(document).on('change', '.mdp-conversion-toggle', function() {
        var $checkbox = $(this);
        var submissionId = $checkbox.data('submission');
        var enabled = $checkbox.is(':checked') ? 1 : 0;
        $.post(config.ajaxUrl, {
            action: 'mdp_toggle_conversion',
            nonce: config.nonce,
            submission_id: submissionId,
            enabled: enabled
        }).done(function() {
            showToast();
        }).fail(function() {
            $checkbox.prop('checked', !enabled);
        });
    });

    $(document).on('blur', '.mdp-conversion-value', function() {
        var $input = $(this);
        var submissionId = $input.data('submission');
        var value = $input.val();
        $.post(config.ajaxUrl, {
            action: 'mdp_save_conversion_value',
            nonce: config.nonce,
            submission_id: submissionId,
            value: value
        }).done(function() {
            showToast();
        });
    });

    $(document).on('change', '.mdp-conversion-time', function() {
        var $input = $(this);
        var submissionId = $input.data('submission');
        var value = $input.val();
        var defaultValue = $input.data('default') || '';
        if (value === defaultValue) {
            value = '';
        }
        $.post(config.ajaxUrl, {
            action: 'mdp_save_conversion_time',
            nonce: config.nonce,
            submission_id: submissionId,
            value: value
        }).done(function() {
            showToast();
        });
    });

    if ($consentSelect.length) {
        updateConsentFields();
        $consentSelect.on('change', updateConsentFields);
    }
});
