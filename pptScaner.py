import os
import sys
import json
import time
import threading
import platform
import subprocess
import webview

SCAN_EXTENSIONS = ('.ppt', '.pptx', 'xwb')
BATCH_SIZE = 5
PATH_UPDATE_MS = 200

IGNORE_DIRS = {
    '$recycle.bin',
    'node_modules',
    '.git',
    'System Volume Information'
}

HTML = """
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
/* 1. 全局盒模型设置，解决 padding 导致溢出的核心 */
* {
    box-sizing: border-box; 
}

body {
    font-family: "Segoe UI", sans-serif;
    margin: 0;
    padding: 20px;
    background: #f8f9fa;
    height: 100vh; /* 占满窗口高度 */
    display: flex;
    flex-direction: column;
    overflow: hidden; /* 2. 禁止 body 滚动，防止双重滚动条 */
}

.header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 15px;
    flex-shrink: 0; /* 防止头部被压缩 */
}

.controls {
    display: flex;
    align-items: center;
    gap: 10px;
}

#path-display {
    font-size: 12px;
    color: #555;
    background: #e9f7ef;
    padding: 6px 10px;
    border-radius: 4px;
    margin-bottom: 10px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    flex-shrink: 0; /* 防止路径显示栏被压缩 */
}

#result-list {
    flex: 1; /* 占据剩余空间 */
    overflow-y: auto; /* 3. 只有列表区域可以滚动 */
    background: white;
    border: 1px solid #dcdcdc;
    border-radius: 6px;
    min-height: 0; /* 4. Flex 嵌套滚动的关键修复 */
}

/* 优化滚动条样式 (可选) */
#result-list::-webkit-scrollbar {
    width: 8px;
}
#result-list::-webkit-scrollbar-track {
    background: #f1f1f1;
    border-radius: 6px;
}
#result-list::-webkit-scrollbar-thumb {
    background: #c1c1c1;
    border-radius: 6px;
}
#result-list::-webkit-scrollbar-thumb:hover {
    background: #a8a8a8;
}

.item {
    display: flex;
    justify-content: space-between;
    padding: 10px;
    border-bottom: 1px solid #eee;
    align-items: center;
    gap: 10px;
}

.item:hover {
    background: #eafaf1;
}

.file-info {
    overflow: hidden;
    flex: 1;
}

.file-name {
    font-weight: 600;
    font-size: 14px;
    overflow: hidden;
    text-overflow: ellipsis;
}

.file-path {
    font-size: 12px;
    color: #888;
    overflow: hidden;
    text-overflow: ellipsis;
}

button {
    padding: 6px 12px;
    cursor: pointer;
    border-radius: 4px;
    border: 1px solid #bbb;
    background: white;
    font-size: 12px;
}

#scan-btn {
    background: #2ecc71;
    color: white;
    border: none;
}

#scan-btn:disabled {
    background: #a9dfbf;
}

.btn-open {
    color: #2ecc71;
    border: 1px solid #2ecc71;
}

.btn-open:hover {
    background: #2ecc71;
    color: white;
}

.btn-delete {
    color: #e74c3c;
    border: 1px solid #e74c3c;
}

.btn-delete:hover {
    background: #e74c3c;
    color: white;
}
</style>
</head>

<body>

<div class="header">
    <h2 style="margin:0;font-size:18px;">PPT 搜索器</h2>
    <div class="controls">
        <span id="status" style="font-size:13px">准备就绪</span>
        <button id="scan-btn" onclick="startScan()">开始扫描</button>
        <button id="stop-btn" onclick="stopScan()" disabled>停止</button>
    </div>
</div>

<div id="path-display">等待指令...</div>
<div id="result-list"></div>

<script>
let count = 0;
const list = document.getElementById('result-list');
const status = document.getElementById('status');
const pathDisp = document.getElementById('path-display');

function startScan() {
    count = 0;
    list.innerHTML = '';
    document.getElementById('scan-btn').disabled = true;
    document.getElementById('stop-btn').disabled = false;
    status.innerText = "正在扫描...";
    pywebview.api.start();
}

function stopScan() {
    pywebview.api.stop();
    status.innerText = "正在停止...";
}

function updatePath(p) {
    pathDisp.innerText = p;
}

function deleteFile(btn, path) {
    if(!confirm("确定要删除这个文件吗？\\n" + path)) return;
    
    pywebview.api.delete_file(path).then(result => {
        if (result.success) {
            btn.closest('.item').remove();
            count--;
            status.innerText = "已找到 " + count + " 个文件";
        } else {
            alert("删除失败：\\n" + result.error);
        }
    });
}

function addBatch(files) {
    const fragment = document.createDocumentFragment();
    files.forEach(f => {
        count++;
        const div = document.createElement('div');
        div.className = 'item';
        div.innerHTML = `
            <div class="file-info">
                <div class="file-name">${f.name}</div>
                <div class="file-path">${f.path}</div>
            </div>
            <button class="btn-open"
                onclick='pywebview.api.open_file(${JSON.stringify(f.path)})'>
                打开
            </button>
            <button class="btn-delete"
                onclick='deleteFile(this, ${JSON.stringify(f.path)})'>
                删除
            </button>
        `;
        fragment.appendChild(div);
    });
    list.appendChild(fragment);
    status.innerText = "已找到 " + count + " 个文件";
}

function onDone(msg) {
    document.getElementById('scan-btn').disabled = false;
    document.getElementById('stop-btn').disabled = true;
    status.innerText = "扫描结束";
    pathDisp.innerText = msg;
}
</script>

</body>
</html>
"""

class Api:
    def __init__(self):
        self._window = None
        self._cancel = False
        self._is_scanning = False

    def set_window(self, win):
        self._window = win

    def open_file(self, path):
        try:
            if platform.system() == 'Windows':
                os.startfile(path)
            elif platform.system() == 'Darwin':
                subprocess.Popen(['open', path])
            else:
                subprocess.Popen(['xdg-open', path])
        except Exception as e:
            print("Open error:", e)

    def delete_file(self, path):
        try:
            if os.path.exists(path):
                os.remove(path)
                return {"success": True}
            else:
                return {"success": False, "error": "文件不存在"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def stop(self):
        self._cancel = True

    def start(self):
        if self._is_scanning:
            return
        self._cancel = False
        self._is_scanning = True
        threading.Thread(target=self._scan_thread, daemon=True).start()

    def _scan_thread(self):
        buffer = []
        last_ui_update = time.time()

        try:
            for drive in self._get_drives():
                for root, dirs, files in os.walk(drive, topdown=True):
                    if self._cancel:
                        break
                    
                    dirs[:] = [d for d in dirs if d.lower() not in IGNORE_DIRS and not d.startswith('.')]

                    now = time.time()
                    if now - last_ui_update > PATH_UPDATE_MS / 1000:
                        self._safe_js(f"updatePath({json.dumps(root)})")
                        last_ui_update = now

                    for f in files:
                        if self._cancel: break
                        if f.lower().endswith(SCAN_EXTENSIONS) and not f.startswith('~$'):
                            buffer.append({'name': f, 'path': os.path.join(root, f)})
                        
                        if len(buffer) >= BATCH_SIZE:
                            self._safe_js(f"addBatch({json.dumps(buffer)})")
                            buffer.clear()
                            time.sleep(0.005)
                if self._cancel: break

            if buffer:
                self._safe_js(f"addBatch({json.dumps(buffer)})")

        finally:
            self._is_scanning = False
            msg = "扫描已停止" if self._cancel else "全盘扫描完成"
            self._safe_js(f"onDone({json.dumps(msg)})")

    def _safe_js(self, code):
        if self._window:
            try:
                self._window.evaluate_js(code)
            except:
                pass

    def _get_drives(self):
        if platform.system() == 'Windows':
            import string
            drives = []
            for d in string.ascii_uppercase:
                drive_path = f"{d}:\\"
                if os.path.exists(drive_path):
                    drives.append(drive_path)
            return drives
        return [os.path.expanduser('~')]

def run():
    api = Api()
    window = webview.create_window("PPT 搜索器", html=HTML, js_api=api, width=850, height=650)
    api.set_window(window)
    webview.start()

if __name__ == '__main__':
    run()