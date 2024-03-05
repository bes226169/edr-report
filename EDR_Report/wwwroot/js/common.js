/**
 * 通知(小鈴鐺)
 */
const notify_timeout = 600000;
/**
 * 檢查登入狀態
 */
const check_login_status_timeout = 10000;
/**
 * 檢查使用者目前是否使用手機瀏覽
 */
const isMobile = () => {
    let hasTouchScreen = false;
    if ("maxTouchPoints" in navigator) {
        hasTouchScreen = navigator.maxTouchPoints > 0;
    } else if ("msMaxTouchPoints" in navigator) {
        hasTouchScreen = navigator.msMaxTouchPoints > 0;
    } else {
        let mQ = window.matchMedia && matchMedia("(pointer:coarse)");
        if (mQ && mQ.media === "(pointer:coarse)") {
            hasTouchScreen = !!mQ.matches;
        } else if ('orientation' in window) {
            hasTouchScreen = true;
        } else {
            let UA = navigator.userAgent;
            hasTouchScreen = (
                /\b(BlackBerry|webOS|iPhone|IEMobile)\b/i.test(UA) ||
                /\b(Android|Windows Phone|iPad|iPod)\b/i.test(UA)
            );
        }
    }
    return hasTouchScreen;
}
const BaseUrl = () => document.baseURI.replace(/(\/)$/, '');
/**
 * 判斷變數是否為 undefined、null、空字串、0、空的object、空的array
 * @param {any} o
 */
const isEmpty = (o) => (o === undefined || o === null || (typeof o === 'number' && o === 0) || (typeof o === 'string' && o === '') || (Array.isArray(o) && o.length === 0) || (typeof o === 'object' && Object.keys(o).length === 0));
/**
 * 移除字串中的HTML Tag
 * @param {string} s 字串
 * @param {boolean} rn 移除換行符號(\r、\n)
 */
const removeHtmlTags = (s, rn = false) => {
    if (s == null) s = '';
    if (typeof s === 'string') s = s.replace(/<([^>]+)>/gi, '');
    else s = s.toString();
    if (rn) s = s.replace(/\r/gi, '').replace(/\n/gi, '');
    return s;
};
const dom = {
    id: n => document.getElementById(n),
    class: n => document.getElementsByClassName(m),
    q: n => document.querySelector(n),
    qa: n => document.querySelectorAll(n),
    /**
     * document.createElement
     * @param {any} t tag name
     * @param {any} c class list
     * @param {any} o innerHTML/append
     */
    new: function (t, c, o) {
        let x = document.createElement(t);
        if (typeof c == 'string') x.classList = c;
        this.add(x, [o]);
        return x;
    },
    /**
     * append多個
     * @param {any} o
     * @param {any} a
     */
    add: function (o, a) {
        if (a != null && typeof a == 'object' && a.length > 0) {
            for (let i in a) {
                if (a[i] == null || ['string', 'number', 'object'].indexOf(typeof a[i]) < 0) continue;
                if (['string', 'number'].indexOf(typeof a[i]) >= 0) {
                    o.innerHTML = a[i];
                } else if (a[i] != null && typeof a[i] == 'object') {
                    o.append(a[i]);
                }
            }
        }
        return o;
    },
    /**
     * 設定多個參數setAttribute
     * @param {any} o
     * @param {any} a
     */
    setAttr: function (o, a) {
        if (a != null && typeof a == 'object') {
            for (let i in a) {
                o.setAttribute(i, a[i]);
            }
        }
        return o;
    }
};
/**
 * 發送通知
 * @param {string} title
 * @param {string} body
 * @param {string} icon
 * @param {function} onclick
 * @param {function} onerror
 */
const bes_notify = (title, body, icon, onclick, onerror) => {
    let n;
    if (Notification.permission === 'granted') {
        let set = {
            body: body,
            tag: 'cpm' + new Date().getTime().toString()
        };
        if (typeof icon == 'string' && icon != '') set['icon'] = icon;
        n = new Notification(title, set);
        if (typeof onclick == 'function') n.onclick = onclick;
        if (typeof onerror == 'function') n.onerror = onerror;
    }
    return n;
};
/**
 * Ajax
 * @param {any} url
 * @param {any} data
 * @param {any} suFn
 * @param {any} erFn
 * @param {any} datatype
 * @param {any} method
 */
const ajax = (url, data = null, suFn = null, erFn = null, datatype = 'JSON', method = 'POST') => {
    if (isEmpty(url) || typeof url != 'string') return;
    $.ajax({
        method: method,
        url: BaseUrl() + url,
        dataType: datatype,
        data: data,
        cache: false,
        success: (e) => {
            if (suFn != null && typeof suFn == 'function') suFn(e);
        },
        error: (e) => {
            console.error(e);
            if (erFn != null && typeof erFn == 'function') erFn(e);
        }
    });
}
/**
 * Ajax連線到WebService
 * @param {any} url
 * @param {any} data
 * @param {any} suFn
 * @param {any} datatype
 * @param {any} method
 */
const api = (url, data = null, suFn = null, datatype = 'JSON', method = 'POST') => {
    if (isEmpty(url) || typeof url != 'string') return;
    $.ajax({
        method: method,
        url: url,
        dataType: datatype,
        contentType: 'application/json; charset=utf-8',
        data: isEmpty(data) ? null : JSON.stringify(data),
        cache: false
    }).done((e) => {
        if (suFn != null && typeof suFn == 'function') suFn(isEmpty(e.d) ? e : JSON.parse(e.d));
    });
}
/**
 * 載入頁面
 * @param {any} url
 * @param {any} data
 * @param {any} suFn
 * @param {any} erFn
 */
const load = (url, data = null, suFn = null, erFn = null, method = 'POST') => {
    if (isEmpty(url) || typeof url != 'string') return;
    ajax(url, data, (e) => {
        if (suFn != null && typeof suFn == 'function') suFn(e);
        window_resize();
        scroll2Top();
    }, erFn, 'HTML', method);
}
/**
 * 下載檔案
 * @param {string} url
 * @param {object} data
 * @param {string} filename 檔案名稱
 * @param {any} suFn 完成時執行的程式(若為null則自動下載)
 * @param {any} erFn
 * @param {string} method
 */
const apidl = (url, data = null, filename = '', suFn = null, erFn = null, method = 'POST') => {
    if (isEmpty(url) || typeof url != 'string') return;
    bes_progress.show();
    let fd = new FormData();
    for (let i in data) {
        fd.append(i, data[i]);
    }
    fetch(BaseUrl() + url, {
        method: method,
        responseType: 'blob',
        body: fd
    }).then((e) => {
        bes_progress.hide();
        if (!e.ok) {
            throw "與伺服器連線時發生錯誤！(api)";
        }
        return e.blob();
    }).then((e) => {
        let b = window.URL.createObjectURL(e);
        if (typeof suFn === 'function') {
            suFn(e);
        } else {
            let a = document.createElement('a');
            a.href = b;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            a.remove();
        }
    }).catch((e) => {
        console.error(e);
        bes_progress.hide();
        if (typeof erFn === 'function') erFn(e);
    });
}
/**
 * 上傳檔案
 * @param {String} url
 * @param {Object} data
 * @param {Function} suFn
 * @param {Function} erFn
 */
const apiul = (url, data, suFn = null, erFn) => {
    if (isEmpty(url) || typeof url != 'string' || isEmpty(data)) return;
    let fd = new FormData();
    for (let i in data) {
        fd.append(i, data[i]);
    }
    fetch(BaseUrl() + url, {
        method: 'POST',
        responseType: 'json',
        body: fd
    }).then(e => {
        if (!e.ok) {
            throw "與伺服器連線時發生錯誤！(api)";
        }
        return e.json();
    }).then(e => {
        if (typeof suFn === 'function') suFn(e);
    }).catch(e => {
        console.error(e);
        if (typeof erFn === 'function') erFn(e);
    });
}
/**
 * 檢查統編格式
 * @param {string} gui_number
 */
const check_tax = (gui_number) => {
    if (gui_number.length !== 8) return false;
    let cx = [1, 2, 1, 2, 1, 2, 4, 1];
    let cnum = gui_number.split('');
    let sum = 0;
    let cc = (num) => {
        let total = num;
        if (total > 9) {
            let s = total.toString();
            total = s.substring(0, 1) * 1 + s.substring(1, 2) * 1;
        }
        return total;
    }
    cnum.forEach((item, index) => {
        if (gui_number.charCodeAt() < 48 || gui_number.charCodeAt() > 57) return;
        sum += cc(item * cx[index]);
    });
    return sum % 10 === 0 || (cnum[6] === '7' && (sum + 1) % 10 === 0);
}
/**
 * 檢查Email格式
 * @param {any} str
 */
const check_email = (str) => {
    return /^\w+((-\w+)|(\.\w+))*\@[A-Za-z0-9]+((\.|-)[A-Za-z0-9]+)*\.[A-Za-z]+$/.test(str);
}
/**
 * 檢查電話號碼
 * @param {any} str
 */
const check_tel = (str) => {
    return /(\d{2,3}-?|\(\d{2,3}\))((\d{3,4}-?\d{4})|\d{6,7})/.test(str);
}
/**
 * 檢查手機號碼
 * @param {any} str
 */
const check_cell = (str) => {
    return /09\d{2}(\d{6}|-\d{6}|-\d{3}-\d{3})/.test(str);
}
/** 
 * dialog:載入視窗(只會顯示一個)
 * 顯示:bes_progress.show()
 * 隱藏:bes_progress.hide()
 * 自動判斷:bes_progress.toggle()
 */
var bes_progress = {
    show: function () {
        this._set();
        this._dlg.dialog('open');
    },
    hide: function () {
        let th = this;
        if (this._dlg != null) {
            let h = function () {
                if (th.isShow()) {
                    th._dlg.dialog('close');
                    setTimeout(h, 500);
                }
            }
            setTimeout(h, 500);
        }
    },
    toggle: function () {
        if (this.isShow()) this.hide();
        else this.show();
    },
    isShow: function () {
        return this._dlg != null;
    },
    _dlg: null,
    _set: function () {
        let th = this;
        if (this._dlg == null) {
            this._dlg = $('<div id="dlg_progress">').append(
                $('<div class="progress">').append(
                    $('<div class="progress-bar bg-info progress-bar-striped progress-bar-animated" role="progressbar" aria-valuenow="1" aria-valuemin="0" aria-valuemax="1" style="width:100%">')
                )
            );
            $('body').append(this._dlg);
            this._dlg.dialog({
                title: '載入中，請稍後...',
                size: 'sm',
                closable: false,
                onclosed: () => {
                    th._dlg.dialog('destroy');
                    th._dlg = null;
                }
            });
        }
    }
};
const salert = {
    info: msg =>
        Swal.fire({
            html: msg,
            icon: 'info',
            timer: 2000,
            timerProgressBar: true
        }),
    success: msg =>
        Swal.fire({
            html: msg,
            icon: 'success',
            timer: 2000,
            timerProgressBar: true
        }),
    warning: msg =>
        Swal.fire({
            html: msg,
            icon: 'warning'
        }),
    error: msg =>
        Swal.fire({
            html: msg,
            icon: 'error'
        }),
    confirm: msg =>
        Swal.fire({
            html: msg,
            icon: 'question',
            showCancelButton: true,
            confirmButtonColor: '#3085d6',
            cancelButtonColor: '#d33',
            confirmButtonText: '確定',
            cancelButtonText: '取消',
            closeOnConfirm: false
        })
}