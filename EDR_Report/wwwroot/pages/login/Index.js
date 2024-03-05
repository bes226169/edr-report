$(() => {
    VueApp({
        data: () => ({
            imgCa: ''
        }),
        mounted() {
            this.reset_captcha_img();
        },
        methods: {
            login() {
                this.$loader.show();
                location.href = '/login/erp_sso';
            },
            reset_captcha_img() {
                this.imgCa = '/login/GetCaptchaImage?ts=' + new Date().getTime();
            }
        }
    }, '#app');
});