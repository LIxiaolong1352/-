// AutoClickerService.java - 无障碍自动点击服务
package com.xiaodi.autoclicker;

import android.accessibilityservice.AccessibilityService;
import android.accessibilityservice.GestureDescription;
import android.content.Intent;
import android.graphics.Path;
import android.graphics.PixelFormat;
import android.os.Handler;
import android.os.Looper;
import android.view.Gravity;
import android.view.LayoutInflater;
import android.view.MotionEvent;
import android.view.View;
import android.view.WindowManager;
import android.view.accessibility.AccessibilityEvent;
import android.widget.Button;
import android.widget.FrameLayout;
import android.widget.SeekBar;
import android.widget.TextView;
import android.widget.Toast;

public class AutoClickerService extends AccessibilityService {
    
    private WindowManager windowManager;
    private View floatingView;
    private FrameLayout touchOverlay;
    private boolean isClicking = false;
    private boolean isRecording = false;
    private Handler handler = new Handler(Looper.getMainLooper());
    private Runnable clickRunnable;
    private int clickSpeed = 10; // 次/秒
    private float targetX = -1, targetY = -1;
    
    @Override
    public void onAccessibilityEvent(AccessibilityEvent event) {
        // 不需要处理事件
    }
    
    @Override
    public void onInterrupt() {
        stopClicking();
    }
    
    @Override
    protected void onServiceConnected() {
        super.onServiceConnected();
        showFloatingWindow();
    }
    
    private void showFloatingWindow() {
        windowManager = (WindowManager) getSystemService(WINDOW_SERVICE);
        
        // 创建悬浮控制面板
        floatingView = LayoutInflater.from(this).inflate(R.layout.floating_control, null);
        
        WindowManager.LayoutParams params = new WindowManager.LayoutParams(
            WindowManager.LayoutParams.WRAP_CONTENT,
            WindowManager.LayoutParams.WRAP_CONTENT,
            WindowManager.LayoutParams.TYPE_ACCESSIBILITY_OVERLAY,
            WindowManager.LayoutParams.FLAG_NOT_FOCUSABLE | WindowManager.LayoutParams.FLAG_NOT_TOUCH_MODAL,
            PixelFormat.TRANSLUCENT
        );
        params.gravity = Gravity.TOP | Gravity.START;
        params.x = 100;
        params.y = 200;
        
        // 设置拖动
        setupDrag(floatingView, params);
        
        // 设置控制按钮
        setupControls(floatingView);
        
        windowManager.addView(floatingView, params);
    }
    
    private void setupDrag(View view, WindowManager.LayoutParams params) {
        view.setOnTouchListener(new View.OnTouchListener() {
            private int initialX, initialY;
            private float initialTouchX, initialTouchY;
            
            @Override
            public boolean onTouch(View v, MotionEvent event) {
                switch (event.getAction()) {
                    case MotionEvent.ACTION_DOWN:
                        initialX = params.x;
                        initialY = params.y;
                        initialTouchX = event.getRawX();
                        initialTouchY = event.getRawY();
                        return true;
                    case MotionEvent.ACTION_MOVE:
                        params.x = initialX + (int) (event.getRawX() - initialTouchX);
                        params.y = initialY + (int) (event.getRawY() - initialTouchY);
                        windowManager.updateViewLayout(floatingView, params);
                        return true;
                }
                return false;
            }
        });
    }
    
    private void setupControls(View view) {
        Button btnRecord = view.findViewById(R.id.btn_record);
        Button btnStart = view.findViewById(R.id.btn_start);
        Button btnStop = view.findViewById(R.id.btn_stop);
        SeekBar seekBar = view.findViewById(R.id.seekbar_speed);
        TextView tvSpeed = view.findViewById(R.id.tv_speed);
        
        // 速度调节
        seekBar.setOnSeekBarChangeListener(new SeekBar.OnSeekBarChangeListener() {
            @Override
            public void onProgressChanged(SeekBar seekBar, int progress, boolean fromUser) {
                clickSpeed = progress + 1;
                tvSpeed.setText(clickSpeed + " 次/秒");
            }
            @Override public void onStartTrackingTouch(SeekBar seekBar) {}
            @Override public void onStopTrackingTouch(SeekBar seekBar) {}
        });
        
        // 录制位置按钮
        btnRecord.setOnClickListener(v -> {
            if (isRecording) {
                hideTouchOverlay();
                btnRecord.setText("📍 录制点击位置");
                isRecording = false;
            } else {
                showTouchOverlay();
                btnRecord.setText("✓ 点击屏幕确定位置");
                isRecording = true;
                Toast.makeText(this, "点击屏幕选择连点位置", Toast.LENGTH_SHORT).show();
            }
        });
        
        // 开始连点
        btnStart.setOnClickListener(v -> {
            if (targetX < 0 || targetY < 0) {
                Toast.makeText(this, "请先录制点击位置", Toast.LENGTH_SHORT).show();
                return;
            }
            startClicking();
        });
        
        // 停止连点
        btnStop.setOnClickListener(v -> stopClicking());
    }
    
    private void showTouchOverlay() {
        touchOverlay = new FrameLayout(this);
        touchOverlay.setBackgroundColor(0x44000000);
        
        WindowManager.LayoutParams params = new WindowManager.LayoutParams(
            WindowManager.LayoutParams.MATCH_PARENT,
            WindowManager.LayoutParams.MATCH_PARENT,
            WindowManager.LayoutParams.TYPE_ACCESSIBILITY_OVERLAY,
            WindowManager.LayoutParams.FLAG_NOT_FOCUSABLE,
            PixelFormat.TRANSLUCENT
        );
        
        touchOverlay.setOnTouchListener((v, event) -> {
            if (event.getAction() == MotionEvent.ACTION_DOWN) {
                targetX = event.getRawX();
                targetY = event.getRawY();
                Toast.makeText(this, "位置已记录: (" + (int)targetX + ", " + (int)targetY + ")", Toast.LENGTH_SHORT).show();
                hideTouchOverlay();
                
                // 更新按钮状态
                Button btnRecord = floatingView.findViewById(R.id.btn_record);
                btnRecord.setText("📍 重新录制位置");
                isRecording = false;
            }
            return true;
        });
        
        windowManager.addView(touchOverlay, params);
    }
    
    private void hideTouchOverlay() {
        if (touchOverlay != null) {
            windowManager.removeView(touchOverlay);
            touchOverlay = null;
        }
    }
    
    private void startClicking() {
        if (isClicking) return;
        
        isClicking = true;
        Toast.makeText(this, "开始连点: " + clickSpeed + "次/秒", Toast.LENGTH_SHORT).show();
        
        long interval = 1000 / clickSpeed;
        
        clickRunnable = new Runnable() {
            @Override
            public void run() {
                if (isClicking) {
                    performClick(targetX, targetY);
                    handler.postDelayed(this, interval);
                }
            }
        };
        
        handler.post(clickRunnable);
    }
    
    private void stopClicking() {
        isClicking = false;
        if (clickRunnable != null) {
            handler.removeCallbacks(clickRunnable);
        }
        Toast.makeText(this, "已停止", Toast.LENGTH_SHORT).show();
    }
    
    private void performClick(float x, float y) {
        Path path = new Path();
        path.moveTo(x, y);
        
        GestureDescription.Builder builder = new GestureDescription.Builder();
        builder.addStroke(new GestureDescription.StrokeDescription(path, 0, 50));
        
        dispatchGesture(builder.build(), null, null);
    }
    
    @Override
    public void onDestroy() {
        super.onDestroy();
        stopClicking();
        hideTouchOverlay();
        if (floatingView != null) {
            windowManager.removeView(floatingView);
        }
    }
}
