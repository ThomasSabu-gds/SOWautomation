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
});
