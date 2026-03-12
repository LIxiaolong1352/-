// MainActivity.java - 主界面
package com.xiaodi.autoclicker;

import android.content.Intent;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.provider.Settings;
import android.widget.Button;
import android.widget.Toast;
import androidx.appcompat.app.AppCompatActivity;

public class MainActivity extends AppCompatActivity {
    
    private static final int REQUEST_OVERLAY_PERMISSION = 1001;
    
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        
        Button btnStart = findViewById(R.id.btn_start_service);
        Button btnGuide = findViewById(R.id.btn_guide);
        
        btnStart.setOnClickListener(v -> checkAndStartService());
        
        btnGuide.setOnClickListener(v -> {
            Toast.makeText(this, 
                "使用步骤:\n1. 开启无障碍服务\n2. 允许悬浮窗权限\n3. 录制点击位置\n4. 开始连点", 
                Toast.LENGTH_LONG).show();
        });
    }
    
    private void checkAndStartService() {
        // 检查悬浮窗权限
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.M) {
            if (!Settings.canDrawOverlays(this)) {
                Intent intent = new Intent(Settings.ACTION_MANAGE_OVERLAY_PERMISSION,
                    Uri.parse("package:" + getPackageName()));
                startActivityForResult(intent, REQUEST_OVERLAY_PERMISSION);
                return;
            }
        }
        
        // 打开无障碍设置
        Intent intent = new Intent(Settings.ACTION_ACCESSIBILITY_SETTINGS);
        startActivity(intent);
        
        Toast.makeText(this, "请找到「小弟连点器」并开启无障碍服务", Toast.LENGTH_LONG).show();
    }
    
    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        if (requestCode == REQUEST_OVERLAY_PERMISSION) {
            if (Build.VERSION.SDK_INT < Build.VERSION_CODES.M || Settings.canDrawOverlays(this)) {
                Intent intent = new Intent(Settings.ACTION_ACCESSIBILITY_SETTINGS);
                startActivity(intent);
            }
        }
    }
}
