module.exports = {
    apps: [{
      name: "okaytool",
      cwd: "/root/klickundfertig/buchhaltung/buchhaltung",
      script: "/root/klickundfertig/buchhaltung/buchhaltung/venv/bin/python",
      args: "-m app.main",
      interpreter: "none",
      env: { PORT: "5003" }
    }]
  };