$(() => {
    VueApp({
        data: () => ({
            dept: [],
            proj: [],
            filterProj: [],
            pagingProj: [],
            pageSize: 10,
            pageNumber: 1,
            sortDescProj: true,
            div: 'ZZZ',
            projName: '',
            showDlg: false,
            openDlg: false,
            cbbDiv: null,
        }),
        mounted() {
            $(window).on('resize', this.win_resize);
            this.win_resize();
        },
        methods: {
            win_resize() {
                let wh = window.innerHeight, ww = window.innerWidth;
                if (ww >= 992) {
                    if (wh >= 760) {
                        this.pageSize = 10;
                    } else if (wh >= 546) {
                        this.pageSize = Math.ceil((wh - 546) / 36) + 3;
                    } else {
                        this.pageSize = 3;
                    }
                } else {
                    if (wh >= 566) {
                        this.pageSize = Math.ceil((wh - 566) / 36) + 5;
                    } else {
                        this.pageSize = 5;
                    }
                }
                if (this.openDlg && this.dept.length > 0) {
                    this.to_page(1);
                }
            },
            show_dlg(btn) {
                btn.currentTarget.blur();
                this.$loader.show();
                this.div = 'ZZZ';
                this.projName = '';
                this.sortDescProj = true;
                this.showDlg = true;
                this.dept = [];
                this.proj = [];
                this.filterProj = [];
                this.pagingProj = [];
                this.$ajax('/home/GetProjectList', null, e => {
                    this.$loader.hide();
                    this.dept = e.deptLi;
                    this.proj = e.projLi;
                    this.$nextTick(() => {
                        this.openDlg = true;
                    });
                }, e => {
                    this.$loader.hide();
                    console.error(e);
                    this.$swal.error('與伺服器連線時發生錯誤！');
                });
            },
            init_dlg() {
                this.cbbDiv = new TomSelect('#area_s', {
                    create: false,
                    selectOnTab: true,
                    maxOptions: null,
                    render: {
                        no_results(d) {
                            return `<div class="no-results">找不到包含 「${d.input}」 的項目...</div>`;
                        }
                    }
                });
                this.cbbDiv.setValue(this.div);
                this.search();
            },
            change_div() {
                this.div = this.cbbDiv.activeOption.dataset.value;
                this.search();
            },
            search(se = true) {
                this.filterProj = this.proj.filter(x =>
                    x.DIV == this.div &&
                    (this.$isEmpty(this.projName) || (!this.$isEmpty(this.projName) && !this.$isEmpty(x.PNAME) && x.PNAME.toLowerCase().indexOf(this.projName.toLowerCase()) >= 0))
                );
                this.pageNumber = 1;
                this.sort(null, se);
            },
            to_page(n) {
                this.pageNumber = n;
                let s = (this.pageNumber - 1) * this.pageSize;
                let e = s + this.pageSize;
                this.pagingProj = this.filterProj.slice(s, e);
            },
            sort(btn, se = true) {
                if (btn != null && btn.currentTarget != null) btn.currentTarget.blur();
                if (se) this.sortDescProj = !this.sortDescProj;
                this.filterProj = this.filterProj.sort((a, b) => {
                    if (this.sortDescProj) {
                        return a.PNAME < b.PNAME ? 1 : -1;
                    } else {
                        return a.PNAME > b.PNAME ? 1 : -1;
                    }
                });
                this.to_page(this.pageNumber);
            },
            change_project(btn, pid) {
                if (btn != null && btn.currentTarget != null) btn.currentTarget.blur();
                this.$loader.show();
                this.$ajax('/home/ChangeProj', {
                    pid: pid
                }, e => {
                    if (e.state) {
                        location.reload();
                    } else {
                        this.$loader.hide();
                        this.$swal.error(e.msg);
                    }
                }, e => {
                    this.$loader.hide();
                });
            }
        }
    }, '#selectProjectApp');
});