var app = {
    isClicking: false,
    clickInterval: null,
    speed: 10,
    floatingBtnShown: false,

    initialize: function() {
        document.addEventListener('deviceready', this.onDeviceReady.bind(this), false);
    },

    onDeviceReady: function() {
        this.setupEventListeners();
        console.log('小弟连点器已就绪');
    },

    setupEventListeners: function() {
        var speedRange = document.getElementById('speedRange');
        var speedValue = document.getElementById('speedValue');
        var toggleBtn = document.getElementById('toggleBtn');
        var floatingBtn = document.getElementById('floatingBtn');

        speedRange.addEventListener('input', function() {
            app.speed = this.value;
            speedValue.textContent = app.speed;
            if (app.isClicking) {
                app.stopClicking();
                app.startClicking();
            }
        });

        toggleBtn.addEventListener('click', function() {
            app.toggleClicker();
        });

        floatingBtn.addEventListener('click', function() {
            app.toggleFloating();
        });
    },

    toggleClicker: function() {
        if (this.floatingBtnShown) {
            this.hideFloatingButton();
        } else {
            this.showFloatingButton();
        }
    },

    showFloatingButton: function() {
        var floatingBtn = document.getElementById('floatingBtn');
        var toggleBtn = document.getElementById('toggleBtn');
        var statusText = document.getElementById('statusText');

        floatingBtn.style.display = 'block';
        toggleBtn.textContent = '隐藏悬浮按钮';
        toggleBtn.classList.remove('btn-start');
        toggleBtn.classList.add('btn-stop');
        statusText.textContent = '状态: 等待点击悬浮按钮';
        statusText.classList.remove('stopped');
        this.floatingBtnShown = true;
    },

    hideFloatingButton: function() {
        var floatingBtn = document.getElementById('floatingBtn');
        var toggleBtn = document.getElementById('toggleBtn');
        var statusText = document.getElementById('statusText');

        floatingBtn.style.display = 'none';
        toggleBtn.textContent = '开始连点';
        toggleBtn.classList.remove('btn-stop');
        toggleBtn.classList.add('btn-start');
        statusText.textContent = '状态: 已停止';
        statusText.classList.add('stopped');
        this.floatingBtnShown = false;
        this.stopClicking();
    },

    toggleFloating: function() {
        if (this.isClicking) {
            this.stopClicking();
        } else {
            this.startClicking();
        }
    },

    startClicking: function() {
        var floatingBtn = document.getElementById('floatingBtn');
        var statusText = document.getElementById('statusText');

        this.isClicking = true;
        floatingBtn.classList.add('active');
        floatingBtn.textContent = '⏹';
        statusText.textContent = '状态: 连点中 (' + this.speed + '次/秒)';

        var interval = 1000 / this.speed;

        this.clickInterval = setInterval(function() {
            // 模拟点击
            if (typeof cordova !== 'undefined' && cordova.plugins && cordova.plugins.autoClick) {
                cordova.plugins.autoClick.click();
            }
        }, interval);
    },

    stopClicking: function() {
        var floatingBtn = document.getElementById('floatingBtn');
        var statusText = document.getElementById('statusText');

        this.isClicking = false;
        clearInterval(this.clickInterval);
        floatingBtn.classList.remove('active');
        floatingBtn.textContent = '⚡';
        statusText.textContent = '状态: 已停止';
        statusText.classList.add('stopped');
    }
};

app.initialize();
