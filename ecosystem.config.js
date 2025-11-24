module.exports = {
    apps: [{
      name: "okaytool-dev",
      cwd: "/root/klickundfertig/buchhaltung/buchhaltung",
      script: "/root/klickundfertig/buchhaltung/buchhaltung/venv/bin/python",
      args: "-m app.main",
      interpreter: "none",
      env: { PORT: "8000" }
    }]
  };