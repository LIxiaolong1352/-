# 小弟连点器 - Cordova项目

## 在线构建APK步骤

### 方法1：使用 Monaca (推荐)
1. 访问 https://monaca.io/
2. 注册账号
3. 点击 "Create New Project"
4. 选择 "Import Project"
5. 上传此ZIP文件
6. 等待构建完成
7. 下载APK

### 方法2：使用 Ionic Appflow
1. 访问 https://ionic.io/appflow
2. 连接GitHub仓库或上传项目
3. 选择Android构建
4. 下载APK

### 方法3：本地构建
```bash
# 安装依赖
npm install -g cordova

# 进入项目目录
cd xiaodi-clicker-cordova

# 添加Android平台
cordova platform add android

# 构建APK
cordova build android

# 输出位置
platforms/android/app/build/outputs/apk/debug/app-debug.apk
```

## 项目结构
```
xiaodi-clicker-cordova/
├── config.xml          # Cordova配置
├── package.json        # Node配置
├── www/                # 网页文件
│   ├── index.html      # 主页面
│   ├── css/
│   │   └── style.css   # 样式
│   ├── js/
│   │   └── app.js      # 逻辑
│   └── img/            # 图片
└── res/                # 资源文件
    └── icon.png        # 应用图标
```

## 注意事项
- 需要Android SDK（API 24+）
- 构建后的APK需要签名才能发布
- 调试版APK可以直接安装测试
