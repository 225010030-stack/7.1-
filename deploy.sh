#!/bin/bash
# 部署 SSC no-ioa + FastAPI 后端到 21.214.139.16:8081
# 把每条远程命令写成单独的 expect 脚本文件，避免 stdin 冲突
set -e

HOST=21.214.139.16
SSH_PORT=36000
APP_PORT=8081
SSH_USER=root
SSH_PASS='Wangxinwei88666'
REMOTE_DIR=/opt/ssc-noioa
LOCAL_DIR=/Users/zhangwenjing/WorkBuddy/2026-07-01-10-08-47/deploy
PY=python3.12
TMP_DIR=/tmp/ssc-deploy-$$
mkdir -p "$TMP_DIR"

SSH_OPTS="-o StrictHostKeyChecking=no -o ConnectTimeout=15"

# run_remote <name> <command>
# 把命令写到 .exp 文件，避免 shell 转义问题
run_remote() {
  local name="$1"
  local cmd="$2"
  local exp_file="$TMP_DIR/$name.exp"
  cat > "$exp_file" <<EOF
#!/usr/bin/expect -f
set timeout 180
spawn ssh $SSH_OPTS -p $SSH_PORT $SSH_USER@$HOST "$cmd"
expect {
  -re {Password:} { send "$SSH_PASS\r"; exp_continue }
  -re {password:} { send "$SSH_PASS\r"; exp_continue }
  eof
}
EOF
  chmod +x "$exp_file"
  "$exp_file"
}

# scp_file <local_path> <remote_path>
scp_file() {
  local local_path="$1"
  local remote_path="$2"
  local exp_file="$TMP_DIR/scp-$(basename "$local_path").exp"
  cat > "$exp_file" <<EOF
#!/usr/bin/expect -f
set timeout 180
spawn scp $SSH_OPTS -P $SSH_PORT "$local_path" "$SSH_USER@$HOST:$remote_path"
expect {
  -re {Password:} { send "$SSH_PASS\r"; exp_continue }
  -re {password:} { send "$SSH_PASS\r"; exp_continue }
  eof
}
EOF
  chmod +x "$exp_file"
  "$exp_file"
}

echo "=== 部署 SSC no-ioa + FastAPI 后端 ==="
echo "目标: $HOST:$APP_PORT (SSH port $SSH_PORT, python $PY)"
echo ""

echo "[1/7] 测试 SSH + Python 版本..."
run_remote step1 "$PY --version; which $PY; df -h /opt | tail -1"
echo ""

echo "[2/7] 停掉 8081 端口现有进程..."
run_remote step2 "pkill -f 'python3.*app.py' 2>/dev/null; pkill -f 'uvicorn.*app:app' 2>/dev/null; sleep 1; ss -tlnp | grep ':$APP_PORT' || echo PORT_${APP_PORT}_FREE"
echo ""

echo "[3/7] 备份旧版..."
BACKUP_TS=$(date +%Y%m%d_%H%M%S)
run_remote step3 "if test -d $REMOTE_DIR; then mv $REMOTE_DIR /opt/ssc-noioa-backup-$BACKUP_TS; echo BACKED_UP; else echo NO_OLD; fi"
echo ""

echo "[4/7] 创建目录..."
run_remote step4 "mkdir -p $REMOTE_DIR/uploads $REMOTE_DIR/outputs && ls -la $REMOTE_DIR"
echo ""

echo "[5/7] 上传 4 个文件..."
scp_file "$LOCAL_DIR/no-ioa.html" "$REMOTE_DIR/"
scp_file "$LOCAL_DIR/app.py" "$REMOTE_DIR/"
scp_file "$LOCAL_DIR/cleaner.py" "$REMOTE_DIR/"
scp_file "$LOCAL_DIR/requirements.txt" "$REMOTE_DIR/"
echo "上传完成"
run_remote step5_verify "ls -la $REMOTE_DIR"
echo ""

echo "[6/7] 安装 Python 依赖（用 $PY）..."
run_remote step6 "cd $REMOTE_DIR && $PY -m pip install --upgrade pip -i https://pypi.tuna.tsinghua.edu.cn/simple 2>&1 | tail -2 && $PY -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple 2>&1 | tail -10 && echo PIP_DONE"
echo ""

echo "[7/7] 启动 uvicorn..."
run_remote step7 "cd $REMOTE_DIR && nohup $PY -m uvicorn app:app --host 0.0.0.0 --port $APP_PORT > $REMOTE_DIR/server.log 2>&1 & echo \$! > $REMOTE_DIR/server.pid && sleep 3 && echo ---HEALTH--- && curl -s http://127.0.0.1:$APP_PORT/health && echo '' && echo ---HTML--- && curl -sI http://127.0.0.1:$APP_PORT/no-ioa.html | head -3 && echo ---LOG--- && tail -5 $REMOTE_DIR/server.log"
echo ""

echo "=== 公网验证 ==="
sleep 2
curl -sI "http://$HOST:$APP_PORT/no-ioa.html" --max-time 10 | head -3
echo "---"
curl -s "http://$HOST:$APP_PORT/health" --max-time 5
echo ""
echo ""
echo "=== 部署完成 ==="
echo "公网访问: http://$HOST:$APP_PORT/no-ioa.html"

# 清理临时文件
rm -rf "$TMP_DIR"
