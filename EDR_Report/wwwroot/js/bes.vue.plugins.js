/**
 * 建立Vue App
 * 這會自動將所有自製的 plugin 及 component 加進去
 * @param {Object} options createApp 內容
 * @param {Element|String} dom 要套用到哪個元素上
 * @param {Function} oth 其它要增加的設定
 * @returns
 */
const VueApp = (options, dom, oth = null) => {
    let app = Vue.createApp(options);
    if (typeof oth == 'function') oth(app);
    app.use({
        install(app) {
            app.component('bs-modal', vue_common_components.bsModal);
            app.component('bs-paging', vue_common_components.bsPaging);
            app.component('bs-offcanvas', vue_common_components.bsOffcanvas);
            app.component('bs-vr', vue_common_components.bsVr);
            app.component('bs-close', vue_common_components.bsClose);
            app.component('bs-progress', vue_common_components.bsProgress);
            app.component('bs-spinner', vue_common_components.bsSpinner);
            app.component('money-input', vue_common_components.moneyInput);
            app.component('paging-info', vue_common_components.pagingInfo);
            app.use(vue_common_plugins.baseUrl);
            app.use(vue_common_plugins.isEmpty);
            app.use(vue_common_plugins.loader);
            app.use(vue_common_plugins.swal);
            app.use(vue_common_plugins.ajax);
        }
    });
    if (dom != null) app.mount(dom);
    return app;
}
const vue_common_plugins = {
    /**
     * 載入遮罩
     * this.$loader.show()
     * this.$loader.show('Loading...')
     * this.$loader.hide()
     */
    loader: {
        install(app) {
            let h = Vue.h;
            let isShow = false;
            let opt = Vue.reactive({
                msg: '載入中，請稍後...',
            });
            let x = Vue.defineComponent(() => {
                return () => {
                    return h(
                        'div',
                        {
                            'class': 'modal fade',
                            'data-bs-backdrop': 'static',
                            'data-bs-keyboard': 'false',
                            'data-bs-focus': 'true'
                        },
                        h(
                            'div',
                            {
                                'class': 'modal-dialog modal-dialog-centered modal-dialog-scrollable modal-sm'
                            },
                            h(
                                'div',
                                {
                                    'class': 'modal-content'
                                },
                                [
                                    h(
                                        'div',
                                        {
                                            'class': 'modal-header bg-primary text-light'
                                        },
                                        h(
                                            'h5',
                                            opt.msg
                                        )
                                    ),
                                    h(
                                        'div',
                                        {
                                            'class': 'modal-body'
                                        },
                                        h(
                                            'div',
                                            {
                                                'class': 'progress'
                                            },
                                            h(
                                                'div',
                                                {
                                                    'class': 'progress-bar bg-info progress-bar-striped progress-bar-animated',
                                                    'role': 'progressbar',
                                                    'aria-valuenow': 1,
                                                    'aria-valuemin': 0,
                                                    'aria-valuemax': 1,
                                                    'style': {
                                                        'width': '100%'
                                                    }
                                                }
                                            )
                                        )
                                    )
                                ]
                            )
                        )
                    );;

                }
            }, {
                props: {
                    msg: 'loading',
                }
            });
            let z = Vue.createApp(x, opt).mount(document.createElement('div'));
            z.$el.addEventListener('shown.bs.modal', () => {
                isShow = true;
            });
            z.$el.addEventListener('hidden.bs.modal', () => {
                isShow = false;
            });
            let bs = new bootstrap.Modal(z.$el);
            app.config.globalProperties.$loader = {
                show(msg) {
                    if (isEmpty(msg)) msg = '處理中，請稍後...';
                    opt.msg = msg;
                    document.body.append(z.$el);
                    bs.show();
                },
                hide() {
                    let check_is_show = () => {
                        if (isShow) {
                            bs.hide();
                        } else {
                            setTimeout(check_is_show, 50);
                        }
                    }
                    check_is_show();
                }
            }
        }
    },
    /**
     * 取得BaseURI
     * this.$baseUrl
     */
    baseUrl: {
        install(app) {
            app.config.globalProperties.$baseUrl = document.baseURI.replace(/(\/)$/, '');
        }
    },
    /**
     * 檢查變數是否為空
     * 數字 0 也會 return true
     * this.$isEmpty(data)
     */
    isEmpty: {
        install(app) {
            /**
             * 檢查變數是否為空
             * 數字 0 也會 return true
             * @param {any} o 要檢查的變數
             * @returns
             */
            app.config.globalProperties.$isEmpty = o => (o === undefined || o === null || (typeof o === 'number' && o === 0) || (typeof o === 'string' && o === '') || (Array.isArray(o) && o.length === 0) || (typeof o === 'object' && Object.keys(o).length === 0));
        }
    },
    notify: {
        install(app) {
            /**
             * 發送Browser通知
             * @param {String} title 標題
             * @param {String} body 內文
             * @param {String} icon 圖示網址
             * @param {Function} onclick 點擊事件
             * @param {Function} onerror 錯誤處理
             * @returns
             */
            app.config.globalProperties.$notify = (title, body, icon, onclick, onerror) => {
                let n;
                if (Notification.permission === 'granted') {
                    let set = {
                        body: body,
                        tag: 'cpm' + new Date().getTime().toString()
                    }
                    if (typeof icon == 'string' && icon != '') set['icon'] = icon;
                    n = new Notification(title, set);
                    if (typeof onclick == 'function') n.onclick = onclick;
                    if (typeof onerror == 'function') n.onerror = onerror;
                }
                return n;
            }
        }
    },
    swal: {
        install(app) {
            app.config.globalProperties.$swal = {
                info: msg => Swal.fire({
                    html: msg,
                    icon: 'info',
                    timer: 2000,
                    timerProgressBar: true
                }),
                success: msg => Swal.fire({
                    html: msg,
                    icon: 'success',
                    timer: 2000,
                    timerProgressBar: true
                }),
                warning: msg => Swal.fire({
                    html: msg,
                    icon: 'warning'
                }),
                error: msg => Swal.fire({
                    html: msg,
                    icon: 'error'
                }),
                confirm: (msg, evt) => {
                    Swal.fire({
                        html: msg,
                        icon: 'question',
                        showCancelButton: true,
                        confirmButtonColor: '#3085d6',
                        cancelButtonColor: '#d33',
                        confirmButtonText: '確定',
                        cancelButtonText: '取消',
                        closeOnConfirm: false
                    }).then((res) => {
                        if (typeof evt === 'function') evt(res);
                    });
                }
            };
        }
    },
    ajax: {
        install(app) {
            /**
             * jQuery Ajax
             * @param {String} url 目標網址
             * @param {Object} data 要傳遞的資料
             * @param {Function} suFn 成功時呼叫的函數
             * @param {Function} erFn 發生錯誤時呼叫的函數
             * @param {String} datatype 資料類型 Default: JSON
             * @param {String} method GET, POST, PUT, DELETE, etc. Default: POST
             * @returns
             */
            app.config.globalProperties.$ajax = (url, data = null, suFn = null, erFn = null, datatype = 'JSON', method = 'POST') => {
                if (isEmpty(url) || typeof url != 'string') return;
                let baseUrl = document.baseURI.replace(/(\/)$/, '');

                $.ajax({
                    method: method,
                    url: baseUrl + url,
                    dataType: datatype,
                    data: data,
                    cache: false,
                    success: e => {
                        if (suFn != null && typeof suFn == 'function') suFn(e);
                    },
                    error: e => {
                        if (erFn != null && typeof erFn == 'function') erFn(e);
                    }
                });
            }
        }
    }
}
const vue_common_components = {
    bsVr: {
        template: `<div class="vr"></div>`
    },
    /**
     * 自動千分位格式化的 textbox
     */
    moneyInput: {
        // modelValue {number} 值(請直接帶入v-model)
        props: ['modelValue', 'min', 'max', 'padZero', 'noDot'],
        data: () => ({
            isFocus: false
        }),
        template: `<input class="text-end" :value="format(modelValue)" v-on:focusin="f($event.target.value)" v-on:focusout="f2($event.target.value)" @input="upd($event.target.value)" />`,
        methods: {
            f(v) {
                this.isFocus = true;
                this.upd(v);
            },
            f2(v) {
                this.isFocus = false;
                this.upd(v);
            },
            upd(v) {
                this.$emit('update:modelValue', this.ch(v.toString().replace(/[^0-9\.]/g, '')));
            },
            ch(v) {
                let o = v;
                v = Number(v);
                if (this.$isEmpty(v) || isNaN(v)) return o;
                let min = Number(this.min);
                let max = Number(this.max);
                if (!this.$isEmpty(this.min) && !isNaN(min) && v < min) v = min;
                if (!this.$isEmpty(this.max) && !isNaN(max) && v > max) v = max;
                let parts = v.toString().split('.');
                let padZero = parseInt(Number(this.padZero));
                if (!this.noDot && parts.length == 1) parts.push('');
                if (!this.noDot && !this.$isEmpty(this.padZero) && !isNaN(padZero) && padZero > 0) {
                    parts[1] = parts[1].padEnd(padZero, '0').substr(0, padZero);
                }
                if (this.noDot && parts.length > 1) {
                    parts.splice(1, 1);
                }
                if (this.noDot) {
                    v = parts[0];
                } else {
                    v = parts[0] + '.' + parts[1];
                }
                return Number(v);
            },
            format(v) {
                let o = v;
                v = this.ch(v);
                let padZero = parseInt(Number(this.padZero));
                if (!this.isFocus && !this.$isEmpty(v)) {
                    let parts = v.toString().split('.');
                    parts[0] = parts[0].replace(/\B(?<!\.\d*)(?=(\d{3})+(?!\d))/g, ",");
                    if (!this.noDot && parts.length == 1) parts.push('');
                    if (!this.noDot && !this.$isEmpty(this.padZero) && !isNaN(padZero) && padZero > 0) {
                        parts[1] = parts[1].padEnd(padZero, '0').substr(0, padZero);
                    }
                    if (this.noDot && parts.length > 1) {
                        parts.splice(1, 1);
                    }
                    if (this.noDot) {
                        v = parts[0];
                    } else {
                        v = parts[0] + '.' + parts[1];
                    }
                }
                if (this.$isEmpty(v)) v = o;
                return v;
            }
        }
    },
    /**
     * 分頁資訊
     */
    pagingInfo: {
        props: [
            'modelValue', // int, 每頁顯示的筆數;要放到v-model;會直接更新
            'count', // int, 查詢後的總筆數
            'total', // int, 總筆數
            'list' // int[], 分頁方式 ex.[10, 20, 50]
        ],
        methods: {
            on_change(e) {
                this.$emit('update:modelValue', e.currentTarget.value);
            }
        },
        template: `<div class="input-group">
    <span class="input-group-text">共 {{count}} 筆資料</span>
    <span v-if="total > count" class="input-group-text">從 {{total}} 筆資料中過濾</span>
    <template v-if="!$isEmpty(list)">
        <span class="input-group-text">每頁顯示</span>
        <select class="form-select" :value="modelValue" v-on:change="on_change">
            <option v-for="pl in list" :value="pl">{{pl}}</option>
        </select>
        <span class="input-group-text">筆</span>
    </template>
    <slot></slot>
</div>`
    },
    /**
     * Bootstrap Pagination
     */
    bsPaging: {
        props: [
            'max', // int, 總頁數
            'now', // int, 現在的頁碼
            'pageChange', // function, 當按下按鈕後要執行的函數;回傳頁碼
            'count', // int, 顯示的按鈕上限
            'hideText', // boolean, 是否隱藏首頁末頁文字
            'hideFirstAndLast', // boolean, 是否隱藏首頁末頁按鈕
            'hideDot', // boolean, 是否隱藏「...」
            'toFirstHtml', // string, 首頁文字
            'toPrevHtml', // string, 上一頁文字
            'toNextHtml', // string, 下一頁文字
            'toLastHtml', // string, 末頁文字
            'noMarginBottom', // boolean, 是否取消margin-bottom
        ],
        data: () => ({
            c: 10,
            s: true,
            lDot: false,
            rDot: false,
        }),
        watch: {
            max: {
                handler() {
                    this.re();
                }
            },
            now: {
                handler() {
                    this.re();
                }
            },
            count: {
                handler() {
                    this.re();
                }
            },
        },
        methods: {
            re() {
                this.s = false;
                this.$nextTick(() => {
                    this.s = true;
                });
            },
            page_range() {
                let co = Number(this.count);
                if (!isNaN(co) && co > 0) this.c = co;
                let rt = [];
                if (this.max <= this.c) {
                    for (let i = 1; i <= this.max; i++) rt.push(i);
                    this.lDot = false;
                    this.rDot = false;
                } else {
                    let r = parseInt(this.c / 2);
                    let b = this.now - r + 1;
                    let c = this.now + r - 1;
                    if (b < 1) {
                        let d = -b + 1;
                        b = 1;
                        if (c + d > this.max) {
                            c = this.max;
                        } else {
                            c += d;
                        }
                    }
                    if (c > this.max) c = this.max;
                    this.lDot = b > r;
                    this.rDot = c < this.max - r;
                    for (let i = b; i <= c; i++) rt.push(i);
                }
                return rt;
            },
            go(btn, n) {
                btn.currentTarget.blur();
                if (typeof this.pageChange == 'function') this.pageChange(n);
            },
            to_first_html() {
                return this.$isEmpty(this.toFirstHtml) ? '⇤' : this.toFirstHtml;
            },
            to_prev_html() {
                return this.$isEmpty(this.toPrevHtml) ? '←' : this.toPrevHtml;
            },
            to_next_html() {
                return this.$isEmpty(this.toNextHtml) ? '→' : this.toNextHtml;
            },
            to_last_html() {
                return this.$isEmpty(this.toLastHtml) ? '⇥' : this.toLastHtml;
            },
        },
        template: `<nav v-if="s">
    <ul :class="{'pagination': true, 'mb-0': Boolean(noMarginBottom)}">
        <li :class="['page-item', now <= 1 ? 'disabled' : '']" v-if="!Boolean(hideFirstAndLast)">
            <a class="page-link" href="javascript:void(0)" v-on:click="e => go(e, 1)" title="首頁">
                <span v-html="to_first_html()"></span>
                <span v-if="!Boolean(hideText)" class="d-none d-inline ms-1">首頁</span>
            </a>
        </li>
        <li :class="['page-item', now <= 1 ? 'disabled' : '']">
            <a class="page-link" href="javascript:void(0)" v-on:click="e => go(e, now - 1)" title="上一頁">
                <span v-html="to_prev_html()"></span>
                <span v-if="!Boolean(hideText)" class="d-none d-inline ms-1">上一頁</span>
            </a>
        </li>
        <li class="page-item" v-if="lDot && !Boolean(hideDot)"><span class="page-link">...</span></li>
        <li :class="['page-item', now == p ? 'active' : '']" v-for="p in page_range()">
            <a class="page-link" href="javascript:void(0)" v-on:click="e => go(e, p)" :title="'第 ' + p + ' 頁'">{{p}}</a>
        </li>
        <li class="page-item" v-if="rDot && !Boolean(hideDot)"><span class="page-link">...</span></li>
        <li :class="['page-item', now >= max ? 'disabled' : '']">
            <a class="page-link" href="javascript:void(0)" v-on:click="e => go(e, now + 1)" title="下一頁">
                <span v-html="to_next_html()"></span>
                <span v-if="!Boolean(hideText)" class="d-none d-inline ms-1">下一頁</span>
            </a>
        </li>
        <li :class="['page-item', now >= max ? 'disabled' : '']" v-if="!Boolean(hideFirstAndLast)">
            <a class="page-link" href="javascript:void(0)" v-on:click="e => go(e, max)" title="末頁">
                <span v-html="to_last_html()"></span>
                <span v-if="!Boolean(hideText)" class="d-none d-inline ms-1">末頁</span>
            </a>
        </li>
        <slot></slot>
    </ul>
</nav>`
    },
    /**
     * Bootstrap Close 按鈕
     */
    bsClose: {
        props: [
            'white' // boolean
        ],
        methods: {
            get_class() {
                return {
                    'btn-close': true,
                    'btn-close-white': Boolean(this.white)
                };
            }
        },
        template: `<button :class="get_class()"><slot></slot></button>`
    },
    /**
     * Bootstrap Spinner
     */
    bsSpinner: {
        props: [
            'type', // string: 'border', 'grow'
            'size', // string
        ],
        methods: {
            get_type() {
                return this.type != null && typeof this.type == 'string' && ['border', 'grow'].indexOf(this.type.toLowerCase()) >= 0 ? ('spinner-' + this.type.toLowerCase()) : 'spinner-border';
            },
            get_size() {
                return {
                    width: this.size,
                    height: this.size
                }
            }
        },
        template: `<div :class="get_type()" :style="get_size()"><slot></slot></div>`
    },
    /**
     * Bootstrap Progress
     */
    bsProgress: {
        props: [
            'now', // number
            'barClass', // string, array, object
            'max', // number
            'min', // number
            'animated', // boolean
            'striped', // boolean
            'showValue' // string: 'percent', 'value'
        ],
        methods: {
            check_set(d) {
                return typeof d == 'boolean' ? d : Boolean(d);
            },
            bar_class() {
                let c = ['progress-bar', this.check_set(this.striped) ? 'progress-bar-striped' : null, this.check_set(this.animated) ? 'progress-bar-animated' : null];
                if (this.barClass != null) {
                    if (typeof this.barClass == 'string') {
                        c.push(...this.barClass.split(' '));
                    } else if (Array.isArray(this.barClass)) {
                        c.push(...this.barClass);
                    } else if (typeof this.barClass == 'object' && Object.keys(this.barClass).length > 0) {
                        for (let k in this.barClass) {
                            if (this.barClass[k] === true) c.push(k);
                        }
                    }
                }
                return c;
            },
            calc() {
                var rt = '0%';
                let max = Number(this.max);
                if (isNaN(max)) return rt;
                let min = Number(this.min);
                if (isNaN(min)) return rt;
                if (max < min || (max === 0 && min === 0)) return rt;
                let now = Number(this.now);
                if (isNaN(now)) return rt;
                return (Math.round((now - min) / (max - min) * 10000) / 100) + '%';
            },
            show_val() {
                return this.showValue != null && typeof this.showValue == 'string' && ['percent', 'value'].indexOf(this.showValue.toLowerCase()) >= 0 ? (
                    this.showValue.toLowerCase() == 'percent' ? this.calc() : this.now
                ) : '';
            },
            bar_style() {
                return {
                    width: this.calc()
                };
            }
        },
        template: `<div class="progress">
    <div :class="bar_class()" :aria-valuemax="max" :aria-valuemin="min" :aria-valuenow="now" :style="bar_style()">{{show_val()}}</div>
</div>`
    },
    /**
     * Bootstrap Offcanvas
     */
    bsOffcanvas: {
        props: [
            'modelValue', // boolean
            'backdrop', // boolean, string: 'static'
            'keyboard', // boolean
            'scroll', // boolean
            'darkMode', // boolean
            'responsive', // string: 'sm', 'md', 'lg', 'xl', 'xxl'
            'placement', // string: 'start', 'left', 'end', 'right', 'top', 'bottom'
            'onShow', // function
            'onShown', // function
            'onHide', // function
            'onHidden', // function
            'onHidePrevented', // function
            'headerClass', // string, array, object
            'bodyClass', // string, array, object
            'placeholder', // boolean
            'placeholderOpacity', // number
        ],
        data: () => ({
            bx: null
        }),
        watch: {
            modelValue: {
                handler(nv) {
                    if (this.check_resize()) return;
                    if (nv) this.bx.show();
                    else this.bx.hide();
                }
            },
            placeholder: {
                handler(nv) {
                    if (nv) {
                        for (let d of this.$el.querySelectorAll('button, input, textarea, select, a')) {
                            d.blur();
                        }
                        this.$el.style.opacity = this.placeholderOpacity == null || (isNaN(Number(this.placeholderOpacity))) ? '.75' : Number(this.placeholderOpacity).toString();
                    } else {
                        this.$el.style.opacity = null;
                    }
                }
            }
        },
        mounted() {
            this.$el.addEventListener('show.bs.offcanvas', this.on_show);
            this.$el.addEventListener('shown.bs.offcanvas', this.on_shown);
            this.$el.addEventListener('hide.bs.offcanvas', this.on_hide);
            this.$el.addEventListener('hidden.bs.offcanvas', this.on_hidden);
            this.$el.addEventListener('hidePrevented.bs.offcanvas', this.on_hide_prevented);
            this.bx = new bootstrap.Offcanvas(this.$el);
            window.addEventListener('resize', e => {
                this.check_resize(e);
            });
        },
        methods: {
            set_is_show(r) {
                this.$emit('update:modelValue', r);
            },
            check_resize() {
                // 防止啟用responsive後的bug
                if (this.responsive == null || typeof this.responsive != 'string') return false;
                switch (this.responsive.toLowerCase()) {
                    case 'sm':
                        if (window.innerWidth >= 576) {
                            return this.set_is_show(false);
                        }
                        break;
                    case 'md':
                        if (window.innerWidth >= 768) {
                            return this.set_is_show(false);
                        }
                        break;
                    case 'lg':
                        if (window.innerWidth >= 992) {
                            return this.set_is_show(false);
                        }
                        break;
                    case 'xl':
                        if (window.innerWidth >= 1200) {
                            return this.set_is_show(false);
                        }
                        break;
                    case 'xxl':
                        if (window.innerWidth >= 1400) {
                            return this.set_is_show(false);
                        }
                        break;
                }
                return false;
            },
            get_set(n) {
                switch (n) {
                    case 'backdrop':
                        return (this.backdrop == null || (typeof this.backdrop == 'string' && this.backdrop == 'static')) ? 'static' : (typeof this.backdrop == 'boolean' ? this.backdrop : Boolean(this.backdrop));
                    case 'keyboard':
                        return typeof this.keyboard == 'boolean' ? this.keyboard : (this.keyboard == null ? false : Boolean(this.keyboard));
                    case 'scroll':
                        return typeof this.scroll == 'boolean' ? this.scroll : (this.scroll == null ? true : Boolean(this.scroll));
                    case 'darkMode':
                        return typeof this.darkMode == 'boolean' ? this.darkMode : (this.darkMode == null ? false : Boolean(this.darkMode));
                    case 'responsive':
                        return typeof this.responsive == 'string' && ['sm', 'md', 'lg', 'xl', 'xxl'].indexOf(this.responsive.toLowerCase()) >= 0 ? ('offcanvas-' + this.responsive) : 'offcanvas';
                    case 'placement':
                        switch (this.placement) {
                            case 'top':
                                return 'offcanvas-top';
                            case 'bottom':
                                return 'offcanvas-bottom';
                            case 'end':
                            case 'right':
                                return 'offcanvas-end';
                            default:
                                return 'offcanvas-start';
                        }
                    default:
                        return null;
                }
            },
            main_class() {
                return [
                    this.get_set('responsive'),
                    this.get_set('placement'),
                    this.get_set('darkMode') ? 'text-bg-dark' : null,
                ];
            },
            common_class(d, def) {
                let c = [def];
                if (d != null) {
                    if (typeof d == 'string') {
                        c.push(...d.split(' '));
                    } else if (Array.isArray(d)) {
                        c.push(...d);
                    } else if (typeof d == 'object' && Object.keys(d).length > 0) {
                        for (let k in d) {
                            if (d[k] === true) c.push(k);
                        }
                    }
                }
                return c;
            },
            header_class() {
                return this.common_class(this.headerClass, 'offcanvas-header');
            },
            body_class() {
                return this.common_class(this.bodyClass, 'offcanvas-body');
            },
            on_show(...args) {
                if (typeof this.onShow == 'function') this.onShow(...args);
            },
            on_shown(...args) {
                if (typeof this.onShown == 'function') this.onShown(...args);
            },
            on_hide(...args) {
                if (typeof this.onHide == 'function') this.onHide(...args);
            },
            on_hidden(...args) {
                if (typeof this.onHidden == 'function') this.onHidden(...args);
            },
            on_hide_prevented(...args) {
                if (typeof this.onHidePrevented == 'function') this.onHidePrevented(...args);
            },
        },
        template: `<div :class="main_class()" :data-bs-backdrop="get_set('backdrop')" :data-bs-keyboard="get_set('keyboard')" :data-bs-scroll="get_set('scroll')">
    <div :class="header_class()"><slot name="header"></slot></div>
    <div :class="body_class()"><slot></slot></div>
</div>`
    },
    /**
     * Bootstrap Modal
     */
    bsModal: {
        props: [
            'isShow', // boolean
            'fade', // boolean
            'backdrop', // boolean, string: 'static'
            'keyboard', // boolean
            'focus', // boolean
            'size', // string: 'sm', 'lg', 'lg'
            'fullscreen', // boolean: true, string: 'sm', 'md', 'lg', 'xl', 'xxl'
            'center', // boolean
            'scrollable', // boolean
            'onShow', // function
            'onShown', // function
            'onHide', // function
            'onHidden', // function
            'onHidePrevented', // function
            'headerClass', // string, array, object
            'bodyClass', // string, array, object
            'footerClass', // string, array, object
            'placeholder', // boolean
            'placeholderOpacity', // number
        ],
        data: () => ({
            bx: null
        }),
        watch: {
            isShow: {
                handler(nv) {
                    if (nv) this.bx.show();
                    else this.bx.hide();
                }
            },
            placeholder: {
                handler(nv) {
                    if (nv) {
                        for (let d of this.$el.querySelectorAll('button, input, textarea, select, a')) {
                            d.blur();
                        }
                        this.$el.style.opacity = this.placeholderOpacity == null || (isNaN(Number(this.placeholderOpacity))) ? '.75' : Number(this.placeholderOpacity).toString();
                    } else {
                        this.$el.style.opacity = null;
                    }
                }
            }
        },
        mounted() {
            this.$el.addEventListener('show.bs.modal', this.on_show);
            this.$el.addEventListener('shown.bs.modal', this.on_shown);
            this.$el.addEventListener('hide.bs.modal', this.on_hide);
            this.$el.addEventListener('hidden.bs.modal', this.on_hidden);
            this.$el.addEventListener('hidePrevented.bs.modal', this.on_hide_prevented);
            this.bx = new bootstrap.Modal(this.$el);
        },
        methods: {
            get_set(n) {
                switch (n) {
                    case 'backdrop':
                        return (this.backdrop == null || (typeof this.backdrop == 'string' && this.backdrop == 'static')) ? 'static' : (typeof this.backdrop == 'boolean' ? this.backdrop : Boolean(this.backdrop));
                    case 'keyboard':
                        return typeof this.keyboard == 'boolean' ? this.keyboard : (this.keyboard == null ? false : Boolean(this.keyboard));
                    case 'focus':
                        return typeof this.focus == 'boolean' ? this.focus : (this.focus == null ? true : Boolean(this.focus));
                    case 'fade':
                        return typeof this.fade == 'boolean' ? this.fade : (this.fade == null ? true : Boolean(this.fade));
                    case 'center':
                        return typeof this.center == 'boolean' ? this.center : (this.center == null ? true : Boolean(this.center));
                    case 'scrollable':
                        return typeof this.scrollable == 'boolean' ? this.scrollable : (this.scrollable == null ? true : Boolean(this.scrollable));
                    case 'size':
                        return typeof this.size == 'string' && ['sm', 'lg', 'xl'].indexOf(this.size.toLowerCase()) >= 0 ? ('modal-' + this.size.toLowerCase()) : null;
                    case 'fullscreen':
                        return ((typeof this.fullscreen == 'boolean' && this.fullscreen) || typeof this.fullscreen == 'string' && this.fullscreen.toLowerCase() == 'always') ? 'modal-fullscreen' :
                            (
                                typeof this.fullscreen == 'string' && ['sm', 'md', 'lg', 'xl', 'xxl'].indexOf(this.fullscreen.toLowerCase()) >= 0 ? 'modal-fullscreen-' + this.fullscreen + '-down' : null
                            );
                    default:
                        return null;
                }
            },
            main_class() {
                return {
                    'modal': true,
                    'fade': this.get_set('fade')
                };
            },
            dialog_class() {
                return [
                    'modal-dialog',
                    this.get_set('center') ? 'modal-dialog-centered' : null,
                    this.get_set('scrollable') ? 'modal-dialog-scrollable' : null,
                    this.get_set('size'),
                    this.get_set('fullscreen'),
                ];
            },
            common_class(d, def) {
                let c = [def];
                if (d != null) {
                    if (typeof d == 'string') {
                        c.push(...d.split(' '));
                    } else if (Array.isArray(d)) {
                        c.push(...d);
                    } else if (typeof d == 'object' && Object.keys(d).length > 0) {
                        for (let k in d) {
                            if (d[k] === true) c.push(k);
                        }
                    }
                }
                return c;
            },
            header_class() {
                return this.common_class(this.headerClass, 'modal-header');
            },
            body_class() {
                return this.common_class(this.bodyClass, 'modal-body');
            },
            footer_class() {
                return this.common_class(this.footerClass, 'modal-footer');
            },
            on_show(...args) {
                if (typeof this.onShow == 'function') this.onShow(...args);
            },
            on_shown(...args) {
                if (typeof this.onShown == 'function') this.onShown(...args);
            },
            on_hide(...args) {
                if (typeof this.onHide == 'function') this.onHide(...args);
            },
            on_hidden(...args) {
                if (typeof this.onHidden == 'function') this.onHidden(...args);
            },
            on_hide_prevented(...args) {
                if (typeof this.onHidePrevented == 'function') this.onHidePrevented(...args);
            },
        },
        template: `<div :class="main_class()" :data-bs-backdrop="get_set('backdrop')" :data-bs-keyboard="get_set('keyboard')" :data-bs-focus="get_set('focus')">
    <div :class="dialog_class()">
        <div class="modal-content">
            <div :class="header_class()"><slot name="header"></slot></div>
            <div :class="body_class()"><slot></slot></div>
            <div :class="footer_class()"><slot name="footer"></slot></div>
        </div>
    </div>
</div>`
    }
}