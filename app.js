const observer = new IntersectionObserver(
    (entries) => {
        entries.forEach((entry) => {
            if (!entry.isIntersecting) return;
            entry.target.classList.add('visible');
            observer.unobserve(entry.target);
        });
    },
    { threshold: 0.18 }
);

document.querySelectorAll('.service-card, .offer-card, .process article, .statement-box').forEach((element) => {
    element.style.opacity = '0';
    element.style.transform = 'translateY(18px)';
    element.style.transition = 'opacity 500ms ease, transform 500ms ease';
    observer.observe(element);
});

document.addEventListener('scroll', () => {
    document.querySelectorAll('.visible').forEach((element) => {
        element.style.opacity = '1';
        element.style.transform = 'translateY(0)';
    });
});

document.dispatchEvent(new Event('scroll'));

document.querySelectorAll('[data-track-contact]').forEach((element) => {
    element.addEventListener('click', () => {
        if (typeof window.fbq !== 'function') return;

        window.fbq('track', 'Contact', {
            placement: element.getAttribute('data-track-contact'),
            channel: 'whatsapp',
            page: 'jefferson-dev-landing'
        });
    });
});
