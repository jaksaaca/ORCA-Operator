// ==========================================
// 1. NAVIGATION & RESPONSIVE UI LOGIC
// ==========================================
const views = {
  schedule: {
    title: "SCHEDULE DISTRIBUTION",
    sub: "",
  },
  datasync: {
    title: "DATA SYNC ENGINE",
    sub: "",
  },
};

// Mobile Sidebar Controls
const sidebar = document.getElementById("sidebar");
const backdrop = document.getElementById("sidebarBackdrop");
const mobileMenuBtn = document.getElementById("mobileMenuBtn");
const closeSidebarBtn = document.getElementById("closeSidebarBtn");

function toggleSidebar() {
  sidebar.classList.toggle("-translate-x-full");
  backdrop.classList.toggle("hidden");
}

mobileMenuBtn.addEventListener("click", toggleSidebar);
closeSidebarBtn.addEventListener("click", toggleSidebar);
backdrop.addEventListener("click", toggleSidebar);

// Tab Switcher
document.querySelectorAll(".menu-item").forEach((btn) => {
  btn.addEventListener("click", function () {
    // Handle Active states
    document
      .querySelectorAll(".menu-item")
      .forEach((b) => b.classList.remove("active"));
    document
      .querySelectorAll(".app-view")
      .forEach((v) => v.classList.add("hidden"));

    this.classList.add("active");
    const target = this.getAttribute("data-target");
    document.getElementById("view-" + target).classList.remove("hidden");

    // Update Headers
    document.getElementById("headerTitle").innerText = views[target].title;
    document.getElementById("headerSubtitle").innerText = views[target].sub;

    // Auto-close sidebar on mobile after clicking a menu item
    if (window.innerWidth < 1024) {
      toggleSidebar();
    }
  });
});

// ==========================================
// 2. SCHEDULE DISTRIBUTION ENGINE
// ==========================================
const schedPalette = [
  { web: "bg-[#BBDEFB]", excel: "FFBBDEFB" },
  { web: "bg-[#C8E6C9]", excel: "FFC8E6C9" },
  { web: "bg-[#FFE0B2]", excel: "FFFFE0B2" },
  { web: "bg-[#E1BEE7]", excel: "FFE1BEE7" },
  { web: "bg-[#FFCDD2]", excel: "FFFFCDD2" },
  { web: "bg-[#FFF59D]", excel: "FFFFF59D" },
];

let schedData = [];

document
  .getElementById("scheduleInput")
  .addEventListener("change", async function (e) {
    const file = e.target.files[0];
    if (!file) return;

    document.getElementById("schedLog1").innerText = `> Selected: ${file.name}`;
    document.getElementById("schedLog2").classList.remove("hidden");

    try {
      schedData = await parseAndSortExcel(file);
      if (schedData.length > 0) {
        document.getElementById("schedLog3").classList.remove("hidden");
        renderSchedPreview(schedData);
        const btn = document.getElementById("schedDownloadBtn");
        btn.removeAttribute("disabled");
        btn.classList.remove("opacity-40", "cursor-not-allowed");
        await triggerSchedDownload();
      }
    } catch (err) {
      alert("Error membaca Excel.");
    }
    e.target.value = "";
  });

async function parseAndSortExcel(file) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());
  const worksheet = workbook.worksheets[0];

  const schedules = [];
  let headers = {};

  worksheet.eachRow((row, rowNum) => {
    if (rowNum === 1) {
      row.eachCell((cell, colNum) => (headers[cell.text.trim()] = colNum));
    } else if (Object.keys(headers).length > 0) {
      const getVal = (name) => {
        const idx = headers[name];
        if (!idx) return "";
        const c = row.getCell(idx);
        return c.value ? (c.value.result || c.value.toString()).trim() : "";
      };
      const lsTime = getVal("LS Time");
      if (!lsTime) return;
      schedules.push({
        "LS Time": lsTime,
        Duration: getVal("Duration") || "2h 0m",
        Brand: getVal("Brand"),
        Platform: getVal("Platform"),
        Host: getVal("Host"),
        Studio: getVal("Studio"),
      });
    }
  });

  schedules.sort((a, b) => {
    const getMins = (t) => {
      const p = (t || "00:00").split("-")[0].trim().split(":");
      return (parseInt(p[0]) || 0) * 60 + (parseInt(p[1]) || 0);
    };
    return getMins(a["LS Time"]) - getMins(b["LS Time"]);
  });

  const uniqueTimes = [...new Set(schedules.map((i) => i["LS Time"]))];
  const map = {};
  uniqueTimes.forEach(
    (t, i) => (map[t] = schedPalette[i % schedPalette.length]),
  );

  return schedules.map((item) => ({ ...item, _color: map[item["LS Time"]] }));
}

function renderSchedPreview(data) {
  const tbody = document.getElementById("schedTbody");
  tbody.innerHTML = "";
  data.forEach((r) => {
    tbody.innerHTML += `
            <tr class="${r._color.web} border-b border-slate-900/10 hover:brightness-95 transition">
                <td class="p-2 border border-slate-900/20 font-bold">${r["LS Time"]}</td>
                <td class="p-2 border border-slate-900/20">${r["Duration"]}</td>
                <td class="p-2 border border-slate-900/20 font-semibold truncate max-w-[150px]" title="${r["Brand"]}">${r["Brand"]}</td>
                <td class="p-2 border border-slate-900/20">${r["Platform"]}</td>
                <td class="p-2 border border-slate-900/20">${r["Host"]}</td>
                <td class="p-2 border border-slate-900/20">${r["Studio"]}</td>
            </tr>`;
  });
  document.getElementById("schedPlaceholder").classList.add("hidden");
  document.getElementById("schedWrapper").classList.remove("hidden");
}

document
  .getElementById("schedDownloadBtn")
  .addEventListener("click", async () => {
    if (schedData.length > 0) await triggerSchedDownload();
  });

async function triggerSchedDownload() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Live Schedule Mapping");
  ws.columns = [
    { header: "LS Time", key: "lsTime", width: 16 },
    { header: "Duration", key: "dur", width: 13 },
    { header: "Brand", key: "brand", width: 32 },
    { header: "Platform", key: "plat", width: 22 },
    { header: "Host", key: "host", width: 18 },
    { header: "Studio", key: "studio", width: 18 },
  ];

  ws.getRow(1).eachCell((c) => {
    c.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFA500" },
    };
    c.font = { bold: true };
    c.alignment = { horizontal: "center" };
    c.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
  });

  schedData.forEach((r) => {
    const row = ws.addRow({
      lsTime: r["LS Time"],
      dur: r["Duration"],
      brand: r["Brand"],
      plat: r["Platform"],
      host: r["Host"],
      studio: r["Studio"],
    });
    row.eachCell((c, i) => {
      c.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: r._color.excel },
      };
      c.alignment = { horizontal: "center", vertical: "middle" };
      c.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      if (i === 1) c.font = { bold: true };
    });
  });

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `ORCA_Schedule_${new Date().toISOString().slice(0, 10)}.xlsx`;
  link.click();
}

// ==========================================
// 3. DATA SYNC ENGINE
// ==========================================
let syncRawData = [];
let syncMappedData = [];

document.getElementById("syncInput").addEventListener("change", function (e) {
  const file = e.target.files[0];
  if (!file) return;

  document.getElementById("syncLog1").innerText = `> Selected: ${file.name}`;
  document.getElementById("syncLog2").classList.remove("hidden");

  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const ws = workbook.Sheets[workbook.SheetNames[0]];
      const sheetArr = XLSX.utils.sheet_to_json(ws, { header: 1 });

      let headerIdx = -1;
      for (let i = 0; i < sheetArr.length; i++) {
        const str = String(sheetArr[i] || []).toLowerCase();
        if (
          str.includes("gmv") ||
          str.includes("impression") ||
          str.includes("views") ||
          str.includes("sold")
        ) {
          headerIdx = i;
          break;
        }
      }

      if (headerIdx === -1) {
        alert("Sistem tidak bisa menemukan baris judul metrik.");
        return;
      }
      document.getElementById("syncLog3").classList.remove("hidden");

      const headers = sheetArr[headerIdx].map((h) =>
        String(h || "")
          .replace(/\n/g, " ")
          .trim()
          .toLowerCase(),
      );
      const dataRows = sheetArr
        .slice(headerIdx + 1)
        .filter((r) => r.length > 0);

      syncRawData = dataRows.map((r) => {
        let obj = {};
        headers.forEach((h, i) => (obj[h] = r[i]));
        return obj;
      });

      if (syncRawData.length > 0) {
        processSyncData();
        const btn = document.getElementById("syncDownloadBtn");
        btn.removeAttribute("disabled");
        btn.classList.remove("opacity-40", "cursor-not-allowed");
        triggerSyncDownload();
      }
    } catch (err) {
      console.error(err);
      alert("Terjadi kesalahan membaca file metrik.");
    }
  };
  reader.readAsArrayBuffer(file);
  e.target.value = "";
});

// Helper Sync
const getV = (r, keys) => {
  for (let k in r) {
    let ck = k.replace(/\(.*?\)/g, "").trim();
    for (let pk of keys) if (ck === pk.toLowerCase()) return r[k] || 0;
  }
  for (let k in r) {
    for (let pk of keys) {
      if (k.includes(pk.toLowerCase())) {
        if (!pk.includes("rate") && k.includes("rate")) continue;
        if (!pk.includes("ratio") && k.includes("ratio")) continue;
        if (!pk.includes("%") && k.includes("%")) continue;
        if (
          pk.includes("view") &&
          !pk.includes("viewers") &&
          k.includes("viewers")
        )
          continue;
        return r[k] || 0;
      }
    }
  }
  return 0;
};
const getT = (r, keys) => {
  for (let k in r) {
    let ck = k.replace(/\(.*?\)/g, "").trim();
    for (let pk of keys) if (ck === pk.toLowerCase()) return r[k] || "-";
  }
  for (let k in r) {
    for (let pk of keys) if (k.includes(pk.toLowerCase())) return r[k] || "-";
  }
  return "-";
};
const clnCur = (v) => {
  if (typeof v === "number") return v;
  if (!v) return 0;
  let n = parseFloat(
    String(v)
      .replace(/[Rp\s.]/gi, "")
      .replace(",", "."),
  );
  return isNaN(n) ? 0 : n;
};

function processSyncData() {
  syncMappedData = syncRawData.map((row) => {
    const pClick = getV(row, ["product click"]);
    const views = getV(row, ["views", "view"]);
    const co = getV(row, ["order paid for", "co"]);

    const ctr = views > 0 ? ((pClick / views) * 100).toFixed(2) + "%" : "0.00%";
    const coRate =
      pClick > 0 ? ((co / pClick) * 100).toFixed(2) + "%" : "0.00%";

    return {
      Title: getT(row, ["livestream", "title", "judul"]),
      "Duration (s)": getV(row, ["duration"]),
      "Start Time": getT(row, [
        "start time",
        "waktu mulai",
        "live start",
        "time",
      ]),
      GMV: clnCur(getV(row, ["attributed gmv", "direct gmv", "revenue"])),
      "Product Impression": getV(row, ["product impression"]),
      "Product Click": pClick,
      CTR: ctr,
      Order: getV(row, ["item sold", "order"]),
      CO: co,
      "CO Rate": coRate,
      Buyer: getV(row, ["customer", "buyer"]),
      AOV: clnCur(getV(row, ["average price", "aov"])),
      Views: views,
      Viewers: "",
      Comment: getV(row, ["comment", "comments"]),
      Likes: getV(row, ["like", "likes"]),
      Follow: getV(row, ["new follower", "new followers", "follow"]),
      "Avg. View Duration (s)": getV(row, [
        "avg. viewing duration per viewer",
        "avg. view duration",
      ]),
      "Engagement Rate": "",
      "Followers Rate": "",
      "ERR (%)": "",
      "Live Impression": getV(row, ["live impression", "live impressions"]),
    };
  });

  const tbody = document.getElementById("syncTbody");
  tbody.innerHTML = "";

  syncMappedData.slice(0, 10).forEach((r) => {
    tbody.innerHTML += `
            <tr class="border-b border-slate-800/50 hover:bg-slate-800/30 transition">
                <td class="p-2 md:p-2.5 border-r border-slate-800/50 truncate max-w-[150px]" title="${r["Title"]}">${r["Title"]}</td>
                <td class="p-2 md:p-2.5 border-r border-slate-800/50">${r["Start Time"]}</td>
                <td class="p-2 md:p-2.5 border-r border-slate-800/50 text-emerald-400 font-semibold">${r["GMV"]}</td>
                <td class="p-2 md:p-2.5 border-r border-slate-800/50">${r["Product Click"]}</td>
                <td class="p-2 md:p-2.5 border-r border-slate-800/50">${r["CTR"]}</td>
                <td class="p-2 md:p-2.5 border-r border-slate-800/50">${r["Order"]}</td>
                <td class="p-2 md:p-2.5">${r["CO Rate"]}</td>
            </tr>`;
  });

  document.getElementById("syncPlaceholder").classList.add("hidden");
  document.getElementById("syncWrapper").classList.remove("hidden");
}

document
  .getElementById("syncDownloadBtn")
  .addEventListener("click", triggerSyncDownload);

function triggerSyncDownload() {
  if (syncMappedData.length === 0) return;
  const ws = XLSX.utils.json_to_sheet(syncMappedData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "ORCA_Formatted");
  XLSX.writeFile(
    wb,
    `ORCA_Sync_${new Date().toISOString().split("T")[0]}.xlsx`,
  );
}
