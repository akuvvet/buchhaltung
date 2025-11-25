// module.exports = {
//     apps: [{
//       name: "okaytool",
//       cwd: "/root/klickundfertig/buchhaltung/buchhaltung",
//       script: "/root/klickundfertig/buchhaltung/buchhaltung/venv/bin/python",
//       args: "-m app.main",
//       interpreter: "none",
//       env: { PORT: "5003" }
//     }]
//   };

// ecosystem.config.js

module.exports = {
  apps: [
    {
      name: "okay-tools", // Eindeutiger Name für den Prozess
      interpreter: "python3",              // Der Interpreter, der verwendet werden soll (hier: python3)
      script: "./app/main.py",             // Die tatsächliche Python-Datei, die ausgeführt werden soll
      // Alternativ, um das Modul zu starten, verwenden Sie stattdessen 'exec_mode' und 'args':
      // exec_mode: "fork",
      // args: ["-m", "app.main"],
      // ABER: PM2 behandelt den 'script'-Pfad meist zuverlässiger, wenn man ihn auf die Hauptdatei setzt.
      
      watch: false,                        // Auf 'true' setzen, wenn PM2 bei Dateiänderungen neu starten soll
      instances: 1,                        // Anzahl der Instanzen (für Single-Threaded-Apps wie diese: 1)
      autorestart: true,                   // PM2 startet die App bei Absturz automatisch neu
      max_memory_restart: '1G',            // Startet neu, wenn der Speicher über 1 GB geht (optional)
      cwd: "/root/klickundfertig/buchhaltung/buchhaltung", // Wichtig: Arbeitsverzeichnis
      env: {
        NODE_ENV: "production",
      }
    },
  ],
};