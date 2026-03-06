document.addEventListener('DOMContentLoaded', function () {
    var currentPath = window.location.pathname.toLowerCase();
    var links = document.querySelectorAll('.ey-nav-links a');

    links.forEach(function (link) {
        var href = (link.getAttribute('href') || '').toLowerCase();
        if (href && currentPath === href) {
            link.classList.add('active');
        } else if (href && href !== '/' && currentPath.indexOf(href) === 0) {
            link.classList.add('active');
        }
    });

    // --- Loading Overlay ---
    var overlay = document.getElementById('loadingOverlay');
    var loadingText = document.getElementById('loadingText');
    var loadingSubtext = document.getElementById('loadingSubtext');

    function showLoading(text, subtext) {
        if (!overlay) return;
        if (loadingText) loadingText.textContent = text || 'Processing...';
        if (loadingSubtext) loadingSubtext.textContent = subtext || 'Please wait while we prepare your content';
        overlay.style.display = 'flex';
        requestAnimationFrame(function () {
            overlay.classList.add('active');
        });
    }

    function hideLoading() {
        if (!overlay) return;
        overlay.classList.remove('active');
        setTimeout(function () { overlay.style.display = 'none'; }, 300);
    }

    // Show loading on page navigation links (Home -> Create SOW, nav links, breadcrumbs)
    document.querySelectorAll(
        '.ey-hero-actions a, .ey-nav-links a, .ey-breadcrumb a, .ey-stepper-step a, .ey-gen-hero-actions .ey-btn-outline'
    ).forEach(function (link) {
        link.addEventListener('click', function (e) {
            var href = link.getAttribute('href') || '';
            if (!href || href === '#' || href.startsWith('javascript:') || link.getAttribute('target') === '_blank') return;
            showLoading('Loading page...', 'Preparing your workspace');
        });
    });

    // Show loading on form submit (Generate SOW button)
    var sowForm = document.getElementById('sowGenerateForm');
    if (sowForm) {
        sowForm.addEventListener('submit', function () {
            showLoading('Generating SOW...', 'Applying your responses to the template');
        });
    }

    // Show loading on Download SOW button click
    document.querySelectorAll('.ey-gen-hero-actions .ey-btn-yellow').forEach(function (btn) {
        var href = (btn.getAttribute('href') || '').toLowerCase();
        if (href.indexOf('/sow/download') !== -1) {
            btn.addEventListener('click', function () {
                showLoading('Preparing download...', 'Building your finalized SOW document');
                // Hide after a delay since download doesn't navigate away
                setTimeout(hideLoading, 4000);
            });
        }
    });

    // Hide loading when navigating back via browser cache (bfcache)
    window.addEventListener('pageshow', function (e) {
        if (e.persisted) hideLoading();
    });
});
