# 🚀 অদম্য প্রেস — Ebook Creator v2.0 Deployment Guide

## What You'll Set Up

This guide walks you through deploying the Ebook Creator as a live web app that you (and your team) can access from anywhere. By the end, you'll have a working URL like `https://ebook.yourdomain.com` or `http://your-server-ip:8501`.

The app requires **WeasyPrint** (system-level PDF engine) and **Node.js** (for DOCX generation), so we need a VPS server — Streamlit Cloud alone won't work because it doesn't support these system dependencies well.

---

## Option A — Deploy on a VPS (Recommended)

**Best for:** Full control, reliable, supports all features including WeasyPrint + Node.js.

**Cost:** ৳400-800/month ($4-8 USD) on DigitalOcean, Hetzner, or Vultr.

### Step 1: Get a VPS Server

Go to any of these and create an account:

- **DigitalOcean** → digitalocean.com → Create Droplet → Ubuntu 24.04, $6/mo (1GB RAM)
- **Hetzner** → hetzner.com → Cloud → Ubuntu 24.04, €3.79/mo (best value)
- **Vultr** → vultr.com → Deploy → Ubuntu 24.04, $6/mo

When creating, choose **Ubuntu 24.04 LTS**, the cheapest plan (1 vCPU, 1GB RAM), and the region closest to Bangladesh (Singapore `sgp1` is best).

After creation, you'll get a **server IP address** and **root password** (or SSH key).

### Step 2: Connect to Your Server

Open your terminal (or PuTTY on Windows) and connect:

```bash
# Replace with your actual server IP
ssh root@YOUR_SERVER_IP
```

It will ask for the password you got from Step 1. Once connected, you'll see a `root@server:~#` prompt.

### Step 3: Install System Dependencies

Copy-paste these commands one at a time. Each block does one thing:

```bash
# Update the system packages (takes ~1 min)
apt update && apt upgrade -y

# Install Python, pip, and WeasyPrint's system dependencies
apt install -y python3 python3-pip python3-venv \
    libpango1.0-dev libpangocairo-1.0-0 libcairo2-dev \
    libgdk-pixbuf2.0-dev libffi-dev shared-mime-info \
    poppler-utils pandoc fonts-noto fonts-noto-cjk

# Install Node.js 20 (needed for DOCX generation)
curl -fsSL https://deb.nodesource.com/setup_20.x | bash -
apt install -y nodejs

# Install the docx npm package globally
npm install -g docx

# Verify everything installed correctly
python3 --version        # Should show Python 3.12+
node --version           # Should show v20.x
weasyprint --version     # Should show WeasyPrint 60+
```

### Step 4: Create the App Directory

```bash
# Create a dedicated user (more secure than running as root)
adduser --disabled-password --gecos "" ebookapp
mkdir -p /home/ebookapp/ebook-creator
cd /home/ebookapp/ebook-creator
```

### Step 5: Upload Your App Files

From your **local computer** (not the server), open a new terminal and upload the files. You need these 3 files:

```
ebook-creator/
├── app.py                 # Main Streamlit app (1090 lines)
├── generate_docx.js       # Node.js DOCX generator
└── requirements.txt       # Python dependencies
```

Upload them using `scp`:

```bash
# Run this from YOUR LOCAL computer, in the folder where you downloaded the zip
scp app.py generate_docx.js requirements.txt root@YOUR_SERVER_IP:/home/ebookapp/ebook-creator/
```

Or if you have the zip file:

```bash
scp odommo-ebook-creator-v2.zip root@YOUR_SERVER_IP:/home/ebookapp/
ssh root@YOUR_SERVER_IP "cd /home/ebookapp && unzip odommo-ebook-creator-v2.zip && mv ebook-creator/* /home/ebookapp/ebook-creator/"
```

### Step 6: Install Python Dependencies

Back on the server:

```bash
cd /home/ebookapp/ebook-creator

# Create a virtual environment (keeps packages isolated)
python3 -m venv venv
source venv/bin/activate

# Install all Python packages
pip install --upgrade pip
pip install -r requirements.txt

# Also install PyMuPDF for PDF preview rendering
pip install PyMuPDF Pillow
```

### Step 7: Test That It Runs

```bash
cd /home/ebookapp/ebook-creator
source venv/bin/activate

# Quick test — should start without errors
streamlit run app.py --server.port 8501 --server.headless true
```

You should see output like:
```
You can now view your Streamlit app in your browser.
URL: http://0.0.0.0:8501
```

Press `Ctrl+C` to stop it. If you see errors, check that all dependencies installed correctly in Step 3.

### Step 8: Set Up as a Background Service

This makes the app start automatically and run forever, even after you close the terminal:

```bash
# Create a systemd service file
cat > /etc/systemd/system/ebook-creator.service << 'EOF'
[Unit]
Description=Odommo Press Ebook Creator
After=network.target

[Service]
Type=simple
User=ebookapp
WorkingDirectory=/home/ebookapp/ebook-creator
ExecStart=/home/ebookapp/ebook-creator/venv/bin/streamlit run app.py \
    --server.port 8501 \
    --server.headless true \
    --server.maxUploadSize 200 \
    --server.enableXsrfProtection true \
    --browser.gatherUsageStats false
Restart=always
RestartSec=5
Environment="NODE_PATH=/usr/lib/node_modules"

[Install]
WantedBy=multi-user.target
EOF

# Fix file ownership
chown -R ebookapp:ebookapp /home/ebookapp/ebook-creator

# Enable and start the service
systemctl daemon-reload
systemctl enable ebook-creator    # auto-start on boot
systemctl start ebook-creator     # start now
systemctl status ebook-creator    # check it's running
```

You should see `Active: active (running)` in green. Your app is now live at `http://YOUR_SERVER_IP:8501`.

### Step 9: Open the Firewall

```bash
# Allow Streamlit port through the firewall
ufw allow 8501/tcp
ufw allow 22/tcp     # Keep SSH access!
ufw enable           # Activate the firewall
```

### Step 10: Access Your App!

Open your browser and go to:

```
http://YOUR_SERVER_IP:8501
```

Login with password: **odommo2025**

The app is now live and accessible from any device!

---

## Step 11 (Optional): Add a Custom Domain + HTTPS

If you want `https://ebook.yourdomain.com` instead of an IP address:

### 11a: Point Your Domain

Go to your domain registrar (Namecheap, GoDaddy, Cloudflare) and add an **A record**:

```
Type: A
Name: ebook          (or @ for root domain)
Value: YOUR_SERVER_IP
TTL: 300
```

Wait 5-10 minutes for DNS to propagate.

### 11b: Install Nginx + SSL

```bash
# Install Nginx (reverse proxy) and Certbot (free SSL)
apt install -y nginx certbot python3-certbot-nginx

# Create Nginx config
cat > /etc/nginx/sites-available/ebook << 'EOF'
server {
    listen 80;
    server_name ebook.yourdomain.com;   # ← Change this!

    location / {
        proxy_pass http://127.0.0.1:8501;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_read_timeout 600s;
        proxy_send_timeout 600s;
        client_max_body_size 200M;
    }

    location /_stcore/stream {
        proxy_pass http://127.0.0.1:8501/_stcore/stream;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";
    }
}
EOF

# Enable the site and get SSL certificate
ln -s /etc/nginx/sites-available/ebook /etc/nginx/sites-enabled/
nginx -t                                                    # test config
systemctl restart nginx

# Get free SSL certificate (follow the prompts)
certbot --nginx -d ebook.yourdomain.com                    # ← Change this!
```

Now your app is live at `https://ebook.yourdomain.com` with HTTPS!

---

## Useful Server Commands

These are commands you'll use to manage the app after deployment:

```bash
# Check if the app is running
systemctl status ebook-creator

# Restart the app (after updating code)
systemctl restart ebook-creator

# View live logs (to debug issues)
journalctl -u ebook-creator -f

# Stop the app
systemctl stop ebook-creator

# Update the app code (upload new files, then restart)
systemctl restart ebook-creator
```

---

## Option B — Deploy on Railway.app (Simpler, More Expensive)

**Best for:** Quick deployment without server management. **Cost:** ~$5-15/mo based on usage.

### Step 1: Install Railway CLI

```bash
npm install -g @railway/cli
railway login
```

### Step 2: Create Project Files

Create a `Dockerfile` in your project folder:

```dockerfile
FROM python:3.12-slim

# Install WeasyPrint system dependencies
RUN apt-get update && apt-get install -y \
    libpango1.0-dev libpangocairo-1.0-0 libcairo2-dev \
    libgdk-pixbuf2.0-dev libffi-dev shared-mime-info \
    poppler-utils pandoc fonts-noto \
    curl && \
    curl -fsSL https://deb.nodesource.com/setup_20.x | bash - && \
    apt-get install -y nodejs && \
    npm install -g docx && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt PyMuPDF Pillow
COPY . .

EXPOSE 8501
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.headless=true"]
```

### Step 3: Deploy

```bash
cd ebook-creator
railway init
railway up
```

Railway will build the Docker image and deploy it. You'll get a URL like `https://ebook-creator-production.up.railway.app`.

---

## Option C — Deploy with Docker (On Any Server)

If you're comfortable with Docker, this works on any machine:

### Step 1: Create Dockerfile

Use the same Dockerfile from Option B above.

### Step 2: Build and Run

```bash
cd ebook-creator

# Build the image (takes ~3-5 minutes first time)
docker build -t odommo-ebook .

# Run the container
docker run -d \
  --name ebook-creator \
  --restart always \
  -p 8501:8501 \
  odommo-ebook

# Check it's running
docker ps
docker logs ebook-creator
```

Access at `http://YOUR_SERVER_IP:8501`.

---

## Troubleshooting

**Problem: "ModuleNotFoundError: No module named 'weasyprint'"**
→ Make sure you activated the virtual environment: `source venv/bin/activate`

**Problem: "WeasyPrint error: cannot load library 'pango'"**
→ Install system dependencies: `apt install libpango1.0-dev libpangocairo-1.0-0`

**Problem: "node: command not found" or DOCX generation fails**
→ Install Node.js: `curl -fsSL https://deb.nodesource.com/setup_20.x | bash - && apt install nodejs`
→ Install docx: `npm install -g docx`
→ Make sure `generate_docx.js` is in the same folder as `app.py`

**Problem: Bengali fonts look wrong in PDF**
→ Install Noto fonts: `apt install fonts-noto fonts-noto-cjk`

**Problem: App is slow to generate PDFs**
→ Normal — WeasyPrint takes ~15-20 seconds for 50 pages. Upgrade to 2GB RAM server for faster performance.

**Problem: Can't access the app from browser**
→ Check firewall: `ufw allow 8501/tcp`
→ Check service: `systemctl status ebook-creator`
→ Check logs: `journalctl -u ebook-creator -f`

---

## File Structure (What You Need to Deploy)

```
ebook-creator/
├── app.py                  # Main Streamlit app (all features)
├── generate_docx.js        # Node.js DOCX generator (design-matched output)
├── requirements.txt        # Python dependencies
└── (optional) Dockerfile   # For Docker/Railway deployment
```

That's it — just 3 files needed!
