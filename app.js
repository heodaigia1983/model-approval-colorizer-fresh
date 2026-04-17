let API = null;
let excelRows = [];

const colorMap = {
  "SYSTEM 2": "#4CAF50",
  "SYSTEM 5": "#F44336",
  "": "#BDBDBD"
};

function log(msg) {
  document.getElementById("log").textContent += msg + "\n";
}

function clearLog() {
  document.getElementById("log").textContent = "";
}

async function initAPI() {
  if (API) return API;

  API = await TrimbleConnectWorkspace.connect(window.parent, (event, data) => {
    console.log("Trimble event:", event, data);
  });

  log("Đã kết nối Trimble API.");
  return API;
}

function readExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = e => {
      try {
        const workbook = XLSX.read(e.target.result, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        resolve(rows);
      } catch (err) {
        reject(err);
      }
    };

    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function normalizeRows(rows) {
  return rows
    .map(r => ({
      guid: String(r.GUID || "").trim(),
      paintCode: String(r["PAINT CODE"] || "").trim()
    }))
    .filter(r => r.guid);
}

function normalizeRuntimeIds(value) {
  if (value === undefined || value === null) return [];
  if (typeof value === "number") return [value];
  if (Array.isArray(value)) return value.flat(Infinity).filter(v => typeof v === "number");
  return [];
}

function uniqueIds(ids) {
  return [...new Set(ids)];
}

function chunkArray(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) {
    out.push(arr.slice(i, i + size));
  }
  return out;
}

async function getLoadedModelId() {
  const api = await initAPI();

  try {
    const viewerObjects = await api.viewer.getObjects();

    if (Array.isArray(viewerObjects) && viewerObjects.length) {
      const modelIds = [...new Set(viewerObjects.map(x => x.modelId).filter(Boolean))];
      if (modelIds.length) {
        log("Loaded modelIds trong viewer: " + modelIds.join(", "));
        return modelIds[0];
      }
    }

    if (
      viewerObjects &&
      Array.isArray(viewerObjects.modelObjectIds) &&
      viewerObjects.modelObjectIds.length
    ) {
      const modelIds = [...new Set(viewerObjects.modelObjectIds.map(x => x.modelId).filter(Boolean))];
      if (modelIds.length) {
        log("Loaded modelIds trong viewer: " + modelIds.join(", "));
        return modelIds[0];
      }
    }
  } catch (err) {
    log("getObjects fallback: " + (err?.message || String(err)));
  }

  const models = await api.viewer.getModels();

  if (!models || !models.length) {
    throw new Error("Không tìm thấy model đang load.");
  }

  log("viewer.getModels(): " + models.map(m => m.id).join(", "));
  return models[0].id;
}

async function resetViewerColorsOnly() {
  const api = await initAPI();

  try {
    await api.viewer.setObjectState(undefined, {
      color: "reset"
    });
    log("Đã reset lớp màu cũ.");
  } catch (err) {
    log("Reset màu fallback: " + (err?.message || String(err)));
  }
}

async function applyColorGroups(api, modelId, groups) {
  for (const color in groups) {
    const ids = uniqueIds(groups[color]);
    if (!ids.length) continue;

    const batches = chunkArray(ids, 1000);

    for (let i = 0; i < batches.length; i++) {
      const batchIds = batches[i];

      await api.viewer.setObjectState(
        {
          modelObjectIds: [
            {
              modelId: modelId,
              objectRuntimeIds: batchIds
            }
          ]
        },
        {
          color: color
        }
      );

      log(`Đã tô ${batchIds.length} object -> ${color} (batch ${i + 1}/${batches.length})`);
    }
  }
}

async function colorByPaintCode() {
  const api = await initAPI();

  if (!excelRows.length) {
    log("Chưa có dữ liệu Excel.");
    return;
  }

  const modelId = await getLoadedModelId();
  const rows = normalizeRows(excelRows);

  log("ModelId: " + modelId);
  log("Bắt đầu đổi GUID -> runtimeId...");

  const guids = rows.map(r => r.guid);

  if (!guids.length) {
    log("Không có GUID hợp lệ trong Excel.");
    return;
  }

  log("Test GUID đầu tiên: " + guids[0]);

  let testRuntimeIds;
  try {
    testRuntimeIds = await api.viewer.convertToObjectRuntimeIds(modelId, [guids[0]]);
    log("Test runtimeIds[0]: " + JSON.stringify(testRuntimeIds));
  } catch (err) {
    log("Lỗi test convert 1 GUID: " + (err?.message || JSON.stringify(err) || String(err)));
    throw err;
  }

  let runtimeIds;
  try {
    runtimeIds = await api.viewer.convertToObjectRuntimeIds(modelId, guids);
  } catch (err) {
    log("Lỗi convert full list: " + (err?.message || JSON.stringify(err) || String(err)));
    throw err;
  }

  const groups = {};
  let matched = 0;
  let unmatched = 0;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const ids = normalizeRuntimeIds(runtimeIds[i]);

    if (!ids.length) {
      unmatched++;
      continue;
    }

    const color = colorMap[row.paintCode] || "#2196F3";

    if (!groups[color]) groups[color] = [];
    groups[color].push(...ids);
    matched += ids.length;
  }

  log("Match: " + matched);
  log("Không match: " + unmatched);

  await resetViewerColorsOnly();
  await applyColorGroups(api, modelId, groups);

  log("Hoàn tất tô màu.");
}

async function saveCurrentView() {
  const api = await initAPI();

  let viewName = String(document.getElementById("viewNameInput").value || "").trim();
  if (!viewName) {
    const now = new Date();
    viewName =
      "Approval View - " +
      now.getFullYear() + "-" +
      String(now.getMonth() + 1).padStart(2, "0") + "-" +
      String(now.getDate()).padStart(2, "0") + " " +
      String(now.getHours()).padStart(2, "0") + ":" +
      String(now.getMinutes()).padStart(2, "0");
  }

  const createdView = await api.view.createView({
    name: viewName,
    description: "Saved from Model Approval Colorizer | Developed by Le Van Thao"
  });

  if (!createdView?.id) {
    throw new Error("Không lưu được view.");
  }

  await api.view.updateView({ id: createdView.id });
  await api.view.selectView(createdView.id);

  log("Đã lưu và mở view: " + (createdView.name || viewName));
}

document.getElementById("readBtn").addEventListener("click", async () => {
  try {
    const file = document.getElementById("fileInput").files[0];
    if (!file) {
      log("Chưa chọn file Excel.");
      return;
    }

    clearLog();
    await initAPI();

    excelRows = await readExcel(file);
    log(`Đọc xong ${excelRows.length} dòng.`);
    log("5 dòng đầu:");

    excelRows.slice(0, 5).forEach(r => {
      log(`${r.GUID} | ${r["PAINT CODE"]}`);
    });

    await colorByPaintCode();
  } catch (err) {
    console.error(err);
    log("Lỗi: " + (err?.message || JSON.stringify(err) || String(err)));
  }
});

document.getElementById("saveViewBtn").addEventListener("click", async () => {
  try {
    await saveCurrentView();
  } catch (err) {
    console.error(err);
    log("Lỗi lưu view: " + (err?.message || JSON.stringify(err) || String(err)));
  }
});

document.getElementById("resetBtn").addEventListener("click", async () => {
  try {
    clearLog();
    await resetViewerColorsOnly();
  } catch (err) {
    console.error(err);
    log("Lỗi reset màu: " + (err?.message || JSON.stringify(err) || String(err)));
  }
});
