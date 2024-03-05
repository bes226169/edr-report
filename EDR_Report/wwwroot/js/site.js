$(() => {
    $('#sidebarToggle').on('click', () => {
        sidebar_toggle();
        return false;
    });
    $('body')
        .on('mouseenter', '.dropdown.dropdown-hover', function () {
            $(this).children('.dropdown-menu').addClass('show').attr('data-bs-popper', 'static');
        })
        .on('mouseleave', '.dropdown.dropdown-hover', function () {
            $(this).children('.dropdown-menu').removeClass('show').removeAttr('data-bs-popper');
        });
    $(window).on('resize', window_resize);
    // 設定
    moment.locale('zh-tw');
    window_resize();
    rebuild_perfectscrollbar();
    $('[data-bs-toggle="tooltip"]').each(function () {
        let tt = bootstrap.Tooltip.getOrCreateInstance(this);
        let hide = () => {
            tt.hide();
        }
        this.removeEventListener('click', hide);
        this.addEventListener('click', hide);
    });
    $('.placeholder-glow').removeClass('placeholder-glow');
    $('.placeholder').removeClass('placeholder');
});
function load_check(url) {
    if (url == null || typeof url != 'string') {
        return;
    }
    window.location.href = url;
}

function window_resize() {
    let wh = window.innerHeight, ww = window.innerWidth;
    let nav = dom.q('#master-nav .nav-header'), footer = dom.id('layoutFooter'), nt = dom.q('#master-nav .nav-tool');
    let main = dom.id('layoutSidenav'), menu = dom.id('layoutSidenav_nav'), content = dom.id('layoutSidenav_content');
    if (ww > 767) {
        content.style.height = menu.style.height = main.style.height = `${wh - nav.offsetHeight - footer.offsetHeight - 19}px`;
        main.style.marginTop = null;
        
    } else {
        content.style.height = menu.style.height = main.style.height = `${wh - nav.offsetHeight - footer.offsetHeight - nt.offsetHeight - 14}px`;
        main.style.marginTop = `${nt.offsetHeight - 5}px`;
    }
}
function scroll2Top() { //(內容)至頂
    $('#layoutSidenav_content')[0].scrollTop = 0;
}
function rebuild_perfectscrollbar() { //重建Scrollbar
    let perfect_scrollbar_set = {
        wheelSpeed: 1,
        suppressScrollX: true,
        minScrollbarLength: 20
    }
    new PerfectScrollbar('#sidenavAccordion', perfect_scrollbar_set);
    new PerfectScrollbar('#layoutSidenav_content', perfect_scrollbar_set);
}
function sidebar_toggle() { //選單toggle
    $('body').toggleClass('sb-sidenav-toggled');
    localStorage.setItem('sb|sidebar-toggle', $('body').hasClass('sb-sidenav-toggled'));
}
function win_open(url, wname) { //開新視窗
    if (!isEmpty(url) && typeof url == 'string') url = $('#hidRedirectCPM').val() + escape(url);
    if ($('body').hasClass('sb-sidenav-toggled')) sidebar_toggle();
    if (wname == "" || wname == null) {
        window.open(url);
    } else {
        window.open(url, wname, height = 200, width = 400);
    }
    return false;
}
function change_proj(pid, url) { //開始切換專案
    bes_progress.show();
    ajax('/home/ChangeProject', {
        pid: pid,
        url: url
    }, (e) => {
        if (e.Code == 0) {
            location.href = '/home/index';
        } else {
            bes_progress.hide();
            Swal.fire({
                title: '系統錯誤',
                text: e.Msg,
                icon: 'error'
            });
        }
    }, (e) => {
        bes_progress.hide();
        Swal.fire({
            title: '系統錯誤',
            text: '與伺服器連線時發生錯誤！',
            icon: 'error'
        });
    });
}
function get_aplist(page, groupkind2) { //管報
    if ($('body').hasClass('sb-sidenav-toggled')) sidebar_toggle();
    load('/home/VSList', {
        pageNumber: page,
        GROUPKIND2: groupkind2,
    }, (e) => {
        $('#renderbody').html(e);
    }, (e) => {
        Swal.fire({
            title: '系統錯誤',
            text: '與伺服器連線時發生錯誤！',
            icon: 'error'
        });
    });
    return void (0);
}