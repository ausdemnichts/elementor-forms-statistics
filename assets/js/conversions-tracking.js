(function() {
    var config = window.mdpConversionTracking || {};
    var consentTool = config.consentTool || 'none';
    var categories = config.categories || {};

    function getParam(name) {
        var params = new URLSearchParams(window.location.search);
        return params.get(name) || '';
    }

    function setCookie(name, value, days) {
        var expires = '';
        if (days) {
            var date = new Date();
            date.setTime(date.getTime() + days * 24 * 60 * 60 * 1000);
            expires = '; expires=' + date.toUTCString();
        }
        document.cookie = name + '=' + encodeURIComponent(value) + expires + '; path=/; SameSite=Lax';
    }

    function hasComplianzConsent(category) {
        if (typeof window.cmplz_has_consent === 'function') {
            return window.cmplz_has_consent(category);
        }
        return false;
    }

    function hasUsercentricsConsent(category) {
        if (window.UC_UI && typeof window.UC_UI.getServicesBaseInfo === 'function') {
            var services = window.UC_UI.getServicesBaseInfo() || [];
            return services.some(function(service) {
                var cat = service.categorySlug || service.category || '';
                var status = service.consentStatus;
                if (typeof status === 'undefined') {
                    status = service.consent;
                }
                if (typeof status === 'undefined') {
                    status = service.status;
                }
                return cat === category && status === true;
            });
        }
        return false;
    }

    function hasCcm19Consent(category) {
        if (window.ccm19 && typeof window.ccm19.getConsent === 'function') {
            return !!window.ccm19.getConsent(category);
        }
        if (window.ccm19 && typeof window.ccm19.getConsents === 'function') {
            var consents = window.ccm19.getConsents();
            return consents && consents[category] === true;
        }
        return false;
    }

    function hasConsent() {
        switch (consentTool) {
            case 'none':
                return true;
            case 'complianz':
                return hasComplianzConsent(categories.complianz || 'marketing');
            case 'usercentrics':
                return hasUsercentricsConsent(categories.usercentrics || 'ADVERTISING');
            case 'ccm19':
                return hasCcm19Consent(categories.ccm19 || 'marketing');
            default:
                return false;
        }
    }

    var gclid = getParam('gclid');
    var gbraid = getParam('gbraid');
    var wbraid = getParam('wbraid');

    if (!gclid && !gbraid && !wbraid) {
        return;
    }

    if (!hasConsent()) {
        return;
    }

    setCookie('mdp_efs_consent', '1', 90);
    if (gclid) {
        setCookie('mdp_gclid', gclid, 90);
    }
    if (gbraid) {
        setCookie('mdp_gbraid', gbraid, 90);
    }
    if (wbraid) {
        setCookie('mdp_wbraid', wbraid, 90);
    }
})();
