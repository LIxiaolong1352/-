# 小弟连点器 - Android自动点击器

## 功能
- 在任何应用/页面上自动点击指定位置
- 可调节点击速度（1-50次/秒）
- 悬浮控制面板，随时开始/停止
- 可视化录制点击位置

## 构建APK

### 方法1：使用 GitHub Actions（推荐）
1. 把代码上传到 GitHub 仓库
2. 启用 GitHub Actions
3. 自动构建APK并下载

### 方法2：使用在线构建平台
1. **Appcircle** (https://appcircle.io) - 免费额度足够
2. **Codemagic** (https://codemagic.io) - Flutter支持好
3. **Bitrise** (https://bitrise.io) - 有免费计划

### 方法3：本地构建（需要配置环境）
```bash
# 安装 Android Studio
# 打开项目，同步Gradle
# Build -> Generate Signed Bundle/APK
```

## 使用方法
1. 安装APK
2. 打开应用，点击「启动连点器」
3. 系统设置中找到「小弟连点器」，开启无障碍服务
4. 允许悬浮窗权限
5. 进入目标页面（抢奶茶小程序）
6. 点击悬浮面板上的「录制点击位置」
7. 点击屏幕上要连点的位置
8. 调节速度，点击「开始」

## 注意事项
- 需要Android 7.0+ (API 24+)
- 必须开启无障碍服务才能点击其他应用
- 部分应用可能有防点击检测
