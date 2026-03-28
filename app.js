/*
 * STL 輪郭切断スライサー「CutStack」v2.0
 * Copyright (c) 2026 Hiroyuki Muramatsu / Shinshu-u
 * https://gijyutu.com/main/
 * Released under the MIT License
 */
document.addEventListener("DOMContentLoaded", () => initApp());

function initApp() {
  const state = { mesh: null, modelName: "", result: null, selectedSheetIndex: 0, selectedSliceIndex: 0, show3d: false, show3dCount: 0 };
  // viewBox状態（mm単位）
  let vbX = 0, vbY = 0, vbW = 450, vbH = 300;

  function applyViewBox() {
    const svg = uiSafe.previewWrap.querySelector("svg");
    if (!svg) return;
    svg.setAttribute("viewBox", `${vbX} ${vbY} ${vbW} ${vbH}`);
    const fitW = state.result?.sheetWidth || vbW;
    const fitH = state.result?.sheetHeight || vbH;
    // ズーム率：fitViewBox(板全体表示)に対する倍率
    const fitVbW = fitW, fitVbH = fitH;
    const zoomPct = Math.round(fitVbW / vbW * 100);
    uiSafe.zoomLabel.textContent = `${zoomPct}%`;
  }

  function resetView() {
    if (!state.result) return;
    const sheetW = state.result.sheetWidth;
    const sheetH = state.result.sheetHeight;
    const wrapW = uiSafe.previewWrap.clientWidth;
    const wrapH = uiSafe.previewWrap.clientHeight;
    // wrap比率に合わせてviewBoxを決定（板全体が収まるよう）
    const wrapAspect = wrapW / wrapH;
    const sheetAspect = sheetW / sheetH;
    if (wrapAspect > sheetAspect) {
      // 高さ基準
      vbH = sheetH;
      vbW = sheetH * wrapAspect;
      vbX = -(vbW - sheetW) / 2;
      vbY = 0;
    } else {
      // 幅基準
      vbW = sheetW;
      vbH = sheetW / wrapAspect;
      vbX = 0;
      vbY = -(vbH - sheetH) / 2;
    }
    applyViewBox();
  }
  const config = { EPS: 1e-6, joinTol: 1.8, repairTol: 1.0, minPolylinePoints: 3, minSegmentLength: 0.1, minOpenPolylineLength: 1.5 };

  const ids = [
    "fileInput","thicknessInput","marginInput","sheetWidthInput","sheetHeightInput",
    "joinTolInput","repairTolInput","minSegLenInput","minOpenLenInput",
    "showLabelsToggle","showCompareToggle","showNextOutlineToggle","showScoreToggle",
    "runButton","saveAllButton","saveAllPdfButton",
    "prevSheetButton","nextSheetButton","prevSliceButton","nextSliceButton",
    "zoomInButton","zoomOutButton","zoomResetButton","zoomLabel","toggle3dButton","canvas3d",
    "modelName","modelSize",
    "statusBox","previewWrap","improvementBox",
    "controls3d","show3dLessButton","show3dMoreButton","show3dLabel",
    "sliceInfo","sliceZ","sliceClosed","sliceOpen","sliceRepaired","sliceNextPart",
    "diagnosticWrapBefore","diagnosticWrap","diagnosticTableBody","scoreTableBody"
  ];
  const ui = Object.fromEntries(ids.map(id => [id, document.getElementById(id)]));
  const missing = Object.entries(ui).filter(([, el]) => !el).map(([k]) => k);
  if (missing.length) console.warn("一部のUI要素が見つかりません（シンプル版では正常）: " + missing.join(", "));

  // null安全なuiラッパー（存在しない要素へのアクセスを無視）
  const uiSafe = new Proxy(ui, {
    get(target, key) {
      const el = target[key];
      if (el) return el;
      // 存在しない要素はダミーオブジェクトを返す
      return {
        textContent: "", innerHTML: "", className: "",
        checked: true, disabled: false, value: "",
        style: { display: "" },
        classList: { add: () => {}, remove: () => {}, toggle: () => {} },
        setAttribute: () => {}, getAttribute: () => null,
        addEventListener: () => {},
        querySelector: () => null,
        getBoundingClientRect: () => ({ width: 0, height: 0, left: 0, top: 0 }),
        offsetWidth: 0, offsetHeight: 0, clientWidth: 0, clientHeight: 0,
      };
    },
    set(target, key, value) { if (target[key]) target[key] = value; return true; }
  });

  wireEvents();
  updateButtons();

  function wireEvents() {
    uiSafe.fileInput.addEventListener("change", onFileChange);
    uiSafe.runButton.addEventListener("click", runSlicePipeline);
    uiSafe.saveAllButton.addEventListener("click", saveAllCutOnlySvgs);
    uiSafe.saveAllPdfButton.addEventListener("click", saveAllCutOnlyPdf);
    uiSafe.showLabelsToggle.addEventListener("change", renderPreview);
    uiSafe.showCompareToggle.addEventListener("change", renderDiagnostics);
    uiSafe.showNextOutlineToggle.addEventListener("change", renderAll);

    // ---- 3D表示切り替え ----
    uiSafe.toggle3dButton.addEventListener("click", () => {
      state.show3d = !state.show3d;
      uiSafe.toggle3dButton.textContent = state.show3d ? "配置プレビュー" : "組み立て完成図";
      uiSafe.toggle3dButton.classList.toggle("active", state.show3d);
      uiSafe.previewWrap.style.display = state.show3d ? "none" : "";
      uiSafe.canvas3d.style.display = state.show3d ? "block" : "none";
      uiSafe.controls3d.style.display = state.show3d ? "flex" : "none";
      uiSafe.zoomInButton.disabled = state.show3d;
      uiSafe.zoomOutButton.disabled = state.show3d;
      uiSafe.zoomResetButton.disabled = state.show3d;
      if (state.show3d) {
        state.show3dCount = state.result?.slicesWithPolylines?.length || 0;
        update3dLabel();
        start3d();
      }
    });

    uiSafe.show3dMoreButton.addEventListener("click", () => {
      const max = state.result?.slicesWithPolylines?.length || 0;
      state.show3dCount = Math.min(max, (state.show3dCount || max) + 1);
      update3dLabel();
    });
    uiSafe.show3dLessButton.addEventListener("click", () => {
      state.show3dCount = Math.max(1, (state.show3dCount || 1) - 1);
      update3dLabel();
    });

    function update3dLabel() {
      const max = state.result?.slicesWithPolylines?.length || 0;
      const n = state.show3dCount || max;
      uiSafe.show3dLabel.textContent = n >= max ? `全${max}部品` : `${n} / ${max}部品`;
    }

    // ---- ズーム・パン（viewBox操作方式） ----
    const ZOOM_STEP = 1.25;

    function zoomViewBox(factor, originXpx, originYpx) {
      const svg = uiSafe.previewWrap.querySelector("svg");
      if (!svg) return;
      const rect = uiSafe.previewWrap.getBoundingClientRect();
      // originがなければ中央
      const ox = originXpx ?? rect.width / 2;
      const oy = originYpx ?? rect.height / 2;
      // px → viewBox座標
      const scaleX = vbW / rect.width;
      const scaleY = vbH / rect.height;
      const mx = vbX + ox * scaleX;
      const my = vbY + oy * scaleY;
      // 新しいviewBoxサイズ
      const newVbW = vbW * factor;
      const newVbH = vbH * factor;
      // マウス位置を固定したままズーム
      vbX = mx - ox * (newVbW / rect.width);
      vbY = my - oy * (newVbH / rect.height);
      vbW = newVbW;
      vbH = newVbH;
      applyViewBox();
    }

    uiSafe.zoomInButton.addEventListener("click",  () => zoomViewBox(1 / ZOOM_STEP));
    uiSafe.zoomOutButton.addEventListener("click", () => zoomViewBox(ZOOM_STEP));
    uiSafe.zoomResetButton.addEventListener("click", () => { resetView(); });

    uiSafe.previewWrap.addEventListener("wheel", e => {
      e.preventDefault();
      const rect = uiSafe.previewWrap.getBoundingClientRect();
      zoomViewBox(
        e.deltaY < 0 ? 1 / ZOOM_STEP : ZOOM_STEP,
        e.clientX - rect.left,
        e.clientY - rect.top
      );
    }, { passive: false });

    // ドラッグでパン
    let dragging = false, dragStartX = 0, dragStartY = 0, vbStartX = 0, vbStartY = 0;
    uiSafe.previewWrap.addEventListener("mousedown", e => {
      if (e.button !== 0) return;
      dragging = true;
      dragStartX = e.clientX;
      dragStartY = e.clientY;
      vbStartX = vbX;
      vbStartY = vbY;
      uiSafe.previewWrap.classList.add("grabbing");
      e.preventDefault();
    });
    window.addEventListener("mousemove", e => {
      if (!dragging) return;
      const rect = uiSafe.previewWrap.getBoundingClientRect();
      const scaleX = vbW / rect.width;
      const scaleY = vbH / rect.height;
      vbX = vbStartX - (e.clientX - dragStartX) * scaleX;
      vbY = vbStartY - (e.clientY - dragStartY) * scaleY;
      applyViewBox();
    });
    window.addEventListener("mouseup", () => {
      if (!dragging) return;
      dragging = false;
      uiSafe.previewWrap.classList.remove("grabbing");
    });
    uiSafe.showScoreToggle.addEventListener("change", renderDiagnostics);
    uiSafe.prevSheetButton.addEventListener("click", () => { if (!state.result) return; state.selectedSheetIndex = Math.max(0, state.selectedSheetIndex - 1); renderAll(); });
    uiSafe.nextSheetButton.addEventListener("click", () => { if (!state.result) return; state.selectedSheetIndex = Math.min(state.result.sheets.length - 1, state.selectedSheetIndex + 1); renderAll(); });    uiSafe.prevSliceButton.addEventListener("click", () => { if (!state.result) return; state.selectedSliceIndex = Math.max(0, state.selectedSliceIndex - 1); renderDiagnostics(); updateButtons(); });
    uiSafe.nextSliceButton.addEventListener("click", () => { if (!state.result) return; state.selectedSliceIndex = Math.min(state.result.closureSummary.sliceChecks.length - 1, state.selectedSliceIndex + 1); renderDiagnostics(); updateButtons(); });
  }

  function syncConfig() {
    config.joinTol = positiveNumber(uiSafe.joinTolInput.value, 1.8);
    config.repairTol = Math.max(0, Number(uiSafe.repairTolInput.value) || 1.0);
    config.minSegmentLength = Math.max(0, Number(uiSafe.minSegLenInput.value) || 0.1);
    config.minOpenPolylineLength = Math.max(0, Number(uiSafe.minOpenLenInput.value) || 1.5);
  }

  async function onFileChange(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      setStatus("STLを読み込んでいます...");
      const mesh = await parseSTL(file);
      state.mesh = mesh;
      state.modelName = file.name.replace(/\.stl$/i, "");
      state.result = null;
      uiSafe.modelName.textContent = state.modelName;
      uiSafe.modelSize.textContent = `${round(mesh.bounds.maxX - mesh.bounds.minX, 2)} × ${round(mesh.bounds.maxY - mesh.bounds.minY, 2)} × ${round(mesh.bounds.maxZ - mesh.bounds.minZ, 2)} mm`;
      // 3D表示が出ていたら配置プレビューに戻す
      if (state.show3d) {
        state.show3d = false;
        if (anim3d) { cancelAnimationFrame(anim3d); anim3d = null; }
        uiSafe.toggle3dButton.textContent = "組み立て完成図";
        uiSafe.toggle3dButton.classList.remove("active");
        uiSafe.canvas3d.style.display = "none";
        uiSafe.controls3d.style.display = "none";
        uiSafe.previewWrap.style.display = "";
        uiSafe.zoomInButton.disabled = false;
        uiSafe.zoomOutButton.disabled = false;
        uiSafe.zoomResetButton.disabled = false;
      }
      clearResultViews();
      setStatus("STLの読み込みが完了しました。");
    } catch (err) {
      setStatus(`STLの読み込みに失敗しました: ${err.message || err}`);
    }
  }

  function runSlicePipeline() {
    if (!state.mesh) { setStatus("先にSTLファイルを読み込んでください。"); return; }
    syncConfig();
    const thickness = positiveNumber(uiSafe.thicknessInput.value, 3);
    const sheetWidth = positiveNumber(uiSafe.sheetWidthInput.value, 450);
    const sheetHeight = positiveNumber(uiSafe.sheetHeightInput.value, 300);
    const margin = Math.max(0, Number(uiSafe.marginInput.value) || 2);

    try {
      setStatus("スライス処理を実行しています...");
      const rawSlices = buildSegments(state.mesh, thickness);
      if (!rawSlices.length) throw new Error("スライス結果が得られませんでした。");

      const slicesBefore = rawSlices.map(slice => {
        const traced = tracePolylinesV7(slice.segments);
        return {
          ...slice,
          polylinesBefore: traced.polylines.map(points => ({ points: cleanPolyline(points), repaired: false, repairType: null, repairSteps: [] })),
          matchRows: traced.matchRows,
        };
      });

      const slicesAfter = rawSlices.map((slice, idx) => {
        const inputPolylines = slicesBefore[idx].polylinesBefore.map(p => p.points);
        const repaired = autoRepairPolylines(inputPolylines, config.repairTol)
          .filter(item => item.points.length >= config.minPolylinePoints)
          .filter(item => {
            const a = analyzePolyline(item);
            return a.closed || polylineLength(item.points) >= config.minOpenPolylineLength;
          });
        return { ...slice, polylines: repaired };
      }).filter(s => s.polylines.length > 0);

      if (!slicesAfter.length) throw new Error("断面輪郭を生成できませんでした。");
      const packed = packSlicesTight(slicesAfter, sheetWidth, sheetHeight, margin);
      if (!packed.sheets.length) throw new Error("板サイズ内に配置できる断面がありませんでした。");

      const closureSummary = summarizeSlicePolylines(slicesAfter);
      const worstSliceIndices = [...closureSummary.sliceChecks].sort((a, b) => b.openCount - a.openCount).slice(0, 3).map(x => x.sliceIndex);

      state.result = {
        rawSlices,
        slicesWithPolylinesBefore: slicesBefore,
        slicesWithPolylines: slicesAfter,
        sheets: packed.sheets,
        cutOnlySvgs: null, // オンデマンド生成（saveCurrentCutOnlySvg等で都度生成）
        closureSummary,
        proposal: buildImprovementProposal(closureSummary),
        sheetWidth,
        sheetHeight,
        worstSliceIndices,
      };

      state.selectedSheetIndex = 0;
      state.selectedSliceIndex = 0;

      setStatus([
        "スライス処理が完了しました。",
        `部品数: ${slicesAfter.length}`,
        `板枚数: ${packed.sheets.length}`,
        `閉曲線: ${closureSummary.totalClosed}`,
        `開放輪郭: ${closureSummary.totalOpen}`,
      ].join("\n"));

      renderAll();
    } catch (err) {
      state.result = null;
      clearResultViews();
      setStatus(`スライス処理に失敗しました: ${err.message || err}`);
    }
  }

  function renderAll() {
    renderSummary();
    if (!state.show3d) renderPreview(true);
    renderDiagnostics();
    updateButtons();
  }

  function renderSummary() {
    if (!state.result) return;
    const p = state.result.proposal;
    uiSafe.improvementBox.className = `improvement-box ${p.level === "critical" ? "critical" : "good"}`;
    uiSafe.improvementBox.innerHTML = `<div><strong>${escapeHtml(p.headline)}</strong></div><div>${p.items.map(i => `<div>・${escapeHtml(i)}</div>`).join("")}</div>`;
  }

  function renderPreview(fit = false) {
    if (!state.result) { uiSafe.previewWrap.innerHTML = ""; return; }
    const sheet = state.result.sheets[state.selectedSheetIndex];
    const sheetW = state.result.sheetWidth;
    const sheetH = state.result.sheetHeight;

    uiSafe.previewWrap.innerHTML = makePreviewSvgString(
      sheet, sheetW, sheetH,
      uiSafe.showLabelsToggle.checked,
      state.result.slicesWithPolylines,
      uiSafe.showNextOutlineToggle.checked
    );
    const svg = uiSafe.previewWrap.querySelector("svg");
    if (!svg) return;

    // SVGをwrap全体に広げる（viewBoxでズーム・パンを制御）
    svg.setAttribute("width", "100%");
    svg.setAttribute("height", "100%");
    svg.style.transform = "";

    if (fit) resetView();
    else applyViewBox();
  }

  function renderDiagnostics() {
    if (!state.result) {
      uiSafe.diagnosticWrapBefore.innerHTML = "";
      uiSafe.diagnosticWrap.innerHTML = "";
      uiSafe.diagnosticTableBody.innerHTML = "";
      uiSafe.scoreTableBody.innerHTML = "";
      uiSafe.sliceInfo.textContent = "-";
      uiSafe.sliceZ.textContent = "-";
      uiSafe.sliceClosed.textContent = "-";
      uiSafe.sliceOpen.textContent = "-";
      uiSafe.sliceRepaired.textContent = "-";
      uiSafe.sliceNextPart.textContent = "-";
      return;
    }

    const sc = state.result.closureSummary.sliceChecks[state.selectedSliceIndex];
    const sd = state.result.slicesWithPolylines.find(s => s.index === sc.sliceIndex);
    const sb = state.result.slicesWithPolylinesBefore.find(s => s.index === sc.sliceIndex);
    const nextSd = getNextSlice(state.result.slicesWithPolylines, sc.sliceIndex);

    uiSafe.sliceInfo.textContent = `S${String(sc.sliceIndex).padStart(2, "0")}${state.result.worstSliceIndices.includes(sc.sliceIndex) ? " ★" : ""}`;
    uiSafe.sliceZ.textContent = `${sc.z} mm`;
    uiSafe.sliceClosed.textContent = String(sc.closedCount);
    uiSafe.sliceOpen.textContent = String(sc.openCount);
    uiSafe.sliceRepaired.textContent = String(sc.checks.filter(c => c.repaired).length);
    uiSafe.sliceNextPart.textContent = nextSd ? `S${String(nextSd.index).padStart(2, "0")} あり（青で重ね表示）` : "なし（最終部品）";

    uiSafe.diagnosticWrapBefore.innerHTML = uiSafe.showCompareToggle.checked ? buildSliceDiagnosticSvgBefore(sb, nextSd, uiSafe.showNextOutlineToggle.checked) : "";
    uiSafe.diagnosticWrap.innerHTML = buildSliceDiagnosticSvg(sc, sd, nextSd, uiSafe.showNextOutlineToggle.checked);

    uiSafe.diagnosticTableBody.innerHTML = sc.checks.map(check => {
      const isNearest = check.repairSteps?.some(s => s.includes("最近傍"));
      const repairLabel = check.repaired
        ? (isNearest ? '<span style="color:#ea580c;font-weight:700">最近傍</span>' : check.repairType === "merge" ? "結合" : "閉鎖")
        : "-";
      return `
        <tr>
          <td>P${String(check.polylineIndex).padStart(2, "0")}</td>
          <td class="${check.closed ? "state-closed" : "state-open"}">${check.closed ? "閉" : "開"}</td>
          <td>${repairLabel}</td>
          <td>${check.pointCount}</td>
          <td>${check.gap == null ? "-" : `${check.gap} mm`}</td>
        </tr>
      `;
    }).join("");

    if (uiSafe.showScoreToggle.checked) {
      const rows = (sb?.matchRows || []).slice(0, 40);
      uiSafe.scoreTableBody.innerHTML = rows.length
        ? rows.map(r => `
          <tr>
            <td>${escapeHtml(r.node)}</td>
            <td>${escapeHtml(r.pair)}</td>
            <td>${r.cost}</td>
            <td>${r.angle}</td>
            <td>${r.distance}</td>
            <td>${r.gap}</td>
          </tr>
        `).join("")
        : `<tr><td colspan="6">候補はありません</td></tr>`;
    } else {
      uiSafe.scoreTableBody.innerHTML = `<tr><td colspan="6">ペア候補表示はOFFです</td></tr>`;
    }
  }

  function updateButtons() {
    const has = !!state.result;
    uiSafe.saveAllButton.disabled = !has;
    uiSafe.saveAllPdfButton.disabled = !has;
    uiSafe.toggle3dButton.disabled = !has;
    uiSafe.prevSheetButton.disabled = !has || state.selectedSheetIndex === 0;
    uiSafe.nextSheetButton.disabled = !has || state.selectedSheetIndex >= state.result.sheets.length - 1;
    uiSafe.prevSliceButton.disabled = !has || state.selectedSliceIndex === 0;
    uiSafe.nextSliceButton.disabled = !has || state.selectedSliceIndex >= state.result.closureSummary.sliceChecks.length - 1;
  }

  function clearResultViews() {
    uiSafe.previewWrap.innerHTML = "";
    uiSafe.diagnosticWrapBefore.innerHTML = "";
    uiSafe.diagnosticWrap.innerHTML = "";
    uiSafe.diagnosticTableBody.innerHTML = "";
    uiSafe.scoreTableBody.innerHTML = "";
    uiSafe.improvementBox.textContent = "診断結果はここに表示されます。";
    updateButtons();
  }

  function computeBounds(vertices) {
    if (!vertices.length) return { minX: 0, minY: 0, minZ: 0, maxX: 0, maxY: 0, maxZ: 0 };
    let minX = Infinity, minY = Infinity, minZ = Infinity, maxX = -Infinity, maxY = -Infinity, maxZ = -Infinity;
    for (const v of vertices) {
      if (v.x < minX) minX = v.x;
      if (v.y < minY) minY = v.y;
      if (v.z < minZ) minZ = v.z;
      if (v.x > maxX) maxX = v.x;
      if (v.y > maxY) maxY = v.y;
      if (v.z > maxZ) maxZ = v.z;
    }
    return { minX, minY, minZ, maxX, maxY, maxZ };
  }

  function boundsOfPoints(points) {
    if (!points.length) return { minX: 0, minY: 0, maxX: 0, maxY: 0, width: 0, height: 0 };
    const xs = points.map(p => p.x);
    const ys = points.map(p => p.y);
    const minX = Math.min(...xs), minY = Math.min(...ys), maxX = Math.max(...xs), maxY = Math.max(...ys);
    return { minX, minY, maxX, maxY, width: maxX - minX, height: maxY - minY };
  }

  function centerTriangles(triangles) {
    const verts = triangles.flat();
    const b = computeBounds(verts);
    const cx = (b.minX + b.maxX) / 2, cy = (b.minY + b.maxY) / 2, cz = (b.minZ + b.maxZ) / 2;
    const centered = triangles.map(tri => tri.map(v => ({ x: v.x - cx, y: v.y - cy, z: v.z - cz })));
    return { triangles: centered, bounds: computeBounds(centered.flat()) };
  }

  function parseBinarySTL(buffer) {
    const view = new DataView(buffer);
    if (buffer.byteLength < 84) throw new Error("ファイルサイズが小さすぎます。");
    const faceCount = view.getUint32(80, true);
    const expected = 84 + faceCount * 50;
    if (expected > buffer.byteLength || faceCount === 0) throw new Error("Binary STL の面数情報が不正です。");
    const triangles = [];
    let offset = 84;
    for (let i = 0; i < faceCount; i++) {
      offset += 12;
      const tri = [];
      for (let j = 0; j < 3; j++) {
        tri.push({ x: view.getFloat32(offset, true), y: view.getFloat32(offset + 4, true), z: view.getFloat32(offset + 8, true) });
        offset += 12;
      }
      offset += 2;
      triangles.push(tri);
    }
    return triangles;
  }

  function parseAsciiSTL(text) {
    const lines = text.split(/\r?\n/);
    const triangles = [];
    let current = [];
    for (const line of lines) {
      const m = line.trim().match(/^vertex\s+([+-]?\d*\.?\d+(?:[eE][+-]?\d+)?)\s+([+-]?\d*\.?\d+(?:[eE][+-]?\d+)?)\s+([+-]?\d*\.?\d+(?:[eE][+-]?\d+)?)$/);
      if (m) {
        current.push({ x: Number(m[1]), y: Number(m[2]), z: Number(m[3]) });
        if (current.length === 3) {
          triangles.push(current);
          current = [];
        }
      }
    }
    if (!triangles.length) throw new Error("ASCII STL として頂点を読み取れませんでした。");
    return triangles;
  }

  async function parseSTL(file) {
    const buffer = await file.arrayBuffer();
    const decoder = new TextDecoder("utf-8");
    const headerText = decoder.decode(buffer.slice(0, Math.min(256, buffer.byteLength)));
    let triangles;
    try {
      const isLikelyAscii = /^\s*solid\b/i.test(headerText) && /facet\s+normal/i.test(headerText);
      triangles = isLikelyAscii ? parseAsciiSTL(decoder.decode(buffer)) : parseBinarySTL(buffer);
    } catch {
      triangles = parseAsciiSTL(decoder.decode(buffer));
    }
    if (!triangles.length) throw new Error("STLの三角形データを取得できませんでした。");
    return centerTriangles(triangles);
  }

  function trianglePlaneIntersection(a, b, c, z) {
    const points = [];
    function edgeIntersect(p1, p2) {
      const d1 = p1.z - z, d2 = p2.z - z;
      if (Math.abs(d1) < config.EPS && Math.abs(d2) < config.EPS) return null;
      if ((d1 > config.EPS && d2 > config.EPS) || (d1 < -config.EPS && d2 < -config.EPS)) return null;
      if (Math.abs(d1 - d2) < config.EPS) return null;
      const t = (z - p1.z) / (p2.z - p1.z);
      if (!Number.isFinite(t) || t < -config.EPS || t > 1 + config.EPS) return null;
      return { x: p1.x + t * (p2.x - p1.x), y: p1.y + t * (p2.y - p1.y), z };
    }
    const candidates = [edgeIntersect(a, b), edgeIntersect(b, c), edgeIntersect(c, a)].filter(Boolean);
    for (const p of candidates) {
      if (!points.some(q => Math.hypot(p.x - q.x, p.y - q.y, p.z - q.z) < config.EPS)) points.push(p);
    }
    return points.length === 2 ? points : null;
  }

  function buildSegments(mesh, thickness) {
    const { triangles, bounds } = mesh;
    const slices = [];
    for (let z = bounds.minZ + thickness / 2; z <= bounds.maxZ + config.EPS; z += thickness) {
      const segments = [];
      for (const tri of triangles) {
        const hit = trianglePlaneIntersection(tri[0], tri[1], tri[2], z);
        if (hit) {
          const a = { x: hit[0].x, y: hit[0].y };
          const b = { x: hit[1].x, y: hit[1].y };
          if (pointDistance(a, b) >= config.minSegmentLength) segments.push([a, b]);
        }
      }
      if (segments.length) slices.push({ index: slices.length + 1, z: round(z, 3), segments });
    }
    return slices;
  }

  function polylineLength(points) {
    let len = 0;
    for (let i = 1; i < points.length; i++) len += pointDistance(points[i - 1], points[i]);
    return len;
  }

  function analyzePolyline(polyline) {
    const points = polyline?.points || [];
    const closed = isClosedLike(points);
    const start = points[0] || null, end = points[points.length - 1] || null;
    const gap = start && end ? Math.hypot(start.x - end.x, start.y - end.y) : null;
    return { closed, repaired: !!polyline?.repaired, repairType: polyline?.repairType || null, repairSteps: polyline?.repairSteps || [], pointCount: points.length, gap: gap == null ? null : round(gap, 4) };
  }

  function clusterPoints(points, tol = config.joinTol) {
    if (!points.length) return [];
    const cellSize = Math.max(tol, config.EPS * 10);
    const nodes = [];
    const buckets = new Map();

    function bucketKey(ix, iy) {
      return `${ix},${iy}`;
    }

    function bucketCoords(point) {
      return {
        ix: Math.round(point.x / cellSize),
        iy: Math.round(point.y / cellSize),
      };
    }

    function addNodeToBucket(node, ix, iy) {
      const key = bucketKey(ix, iy);
      if (!buckets.has(key)) buckets.set(key, []);
      buckets.get(key).push(node.id);
      node.bucketKey = key;
    }

    for (const point of points) {
      const base = bucketCoords(point);
      let matched = null;
      let bestDistance = Infinity;

      for (let dx = -1; dx <= 1; dx++) {
        for (let dy = -1; dy <= 1; dy++) {
          const ids = buckets.get(bucketKey(base.ix + dx, base.iy + dy)) || [];
          for (const nodeId of ids) {
            const node = nodes[nodeId];
            const dist = pointDistance(point, node);
            if (dist <= tol && dist < bestDistance) {
              matched = node;
              bestDistance = dist;
            }
          }
        }
      }

      if (!matched) {
        const node = { id: nodes.length, x: point.x, y: point.y, count: 1, bucketKey: null };
        nodes.push(node);
        point.nodeId = node.id;
        addNodeToBucket(node, base.ix, base.iy);
        continue;
      }

      if (matched.bucketKey && buckets.has(matched.bucketKey)) {
        const arr = buckets.get(matched.bucketKey);
        const idx = arr.indexOf(matched.id);
        if (idx >= 0) arr.splice(idx, 1);
        if (!arr.length) buckets.delete(matched.bucketKey);
      }

      matched.x = (matched.x * matched.count + point.x) / (matched.count + 1);
      matched.y = (matched.y * matched.count + point.y) / (matched.count + 1);
      matched.count += 1;
      point.nodeId = matched.id;

      const nextBucket = bucketCoords(matched);
      addNodeToBucket(matched, nextBucket.ix, nextBucket.iy);
    }

    return nodes;
  }

  function angleBetween(a, b) { return Math.atan2(b.y - a.y, b.x - a.x); }
  function normalizeAngleDelta(a, b) {
    let d = Math.abs(a - b);
    while (d > Math.PI) d = Math.abs(d - 2 * Math.PI);
    return d;
  }
  function otherNode(edge, nodeId) { return edge.a === nodeId ? edge.b : edge.a; }

  function pairCostAtNode(nodeId, eid1, eid2, nodes, edges) {
    const o1 = otherNode(edges[eid1], nodeId);
    const o2 = otherNode(edges[eid2], nodeId);
    const angle = Math.abs(Math.PI - normalizeAngleDelta(angleBetween(nodes[nodeId], nodes[o1]), angleBetween(nodes[nodeId], nodes[o2])));
    const dist = (pointDistance(nodes[nodeId], nodes[o1]) + pointDistance(nodes[nodeId], nodes[o2])) * 0.08;
    const gap = pointDistance(nodes[o1], nodes[o2]) * 0.03;
    const total = angle * 28 + dist + gap;
    return { total: round(total, 3), angle: round(angle * 28, 3), distance: round(dist, 3), gap: round(gap, 3) };
  }

  function bestPairingAllowOdd(edgeIds, nodeId, nodes, edges) {
    const memo = new Map();
    function rec(ids) {
      const key = ids.join(",");
      if (memo.has(key)) return memo.get(key);
      if (ids.length === 0) return { cost: 0, pairs: [], unpaired: [] };
      if (ids.length === 1) return { cost: 8, pairs: [], unpaired: [ids[0]] };

      let best = { cost: Infinity, pairs: [], unpaired: [] };
      const first = ids[0];

      const skipRest = ids.slice(1);
      const skipped = rec(skipRest);
      if (skipped.cost + 8 < best.cost) {
        best = { cost: skipped.cost + 8, pairs: skipped.pairs, unpaired: [first, ...skipped.unpaired] };
      }

      for (let i = 1; i < ids.length; i++) {
        const second = ids[i];
        const rest = ids.filter((_, idx) => idx !== 0 && idx !== i);
        const c = pairCostAtNode(nodeId, first, second, nodes, edges);
        const sub = rec(rest);
        const total = c.total + sub.cost;
        if (total < best.cost) {
          best = { cost: total, pairs: [{ a: first, b: second, ...c }, ...sub.pairs], unpaired: sub.unpaired };
        }
      }
      memo.set(key, best);
      return best;
    }
    return rec([...edgeIds].sort((a, b) => a - b));
  }

  function tracePolylinesV7(segments) {
    const filtered = segments.filter(seg => pointDistance(seg[0], seg[1]) >= config.minSegmentLength);
    if (!filtered.length) return { polylines: [], matchRows: [] };

    // ---- 1. 点のクラスタリング ----
    // STL浮動小数点誤差を吸収できる最小限の許容値を使う。
    // joinTol（1mm超）を使うと近傍の別点が誤統合されグラフが切断される。
    // 0.05mm は STL 精度として十分小さく、浮動小数点誤差は吸収できる。
    const CLUSTER_TOL = 0.05;
    const rawPoints = filtered.flatMap(seg => [
      { x: seg[0].x, y: seg[0].y, nodeId: -1 },
      { x: seg[1].x, y: seg[1].y, nodeId: -1 },
    ]);
    const nodes = clusterPoints(rawPoints, CLUSTER_TOL);

    // ---- 2. エッジ構築（同一ノード間は1本のみ） ----
    const edges = [];
    const edgeSeen = new Set();
    for (let i = 0; i < filtered.length; i++) {
      const pA = rawPoints[i * 2], pB = rawPoints[i * 2 + 1];
      if (pA.nodeId === pB.nodeId) continue;
      const key = pA.nodeId < pB.nodeId ? `${pA.nodeId}:${pB.nodeId}` : `${pB.nodeId}:${pA.nodeId}`;
      if (edgeSeen.has(key)) continue;
      edgeSeen.add(key);
      edges.push({ id: edges.length, a: pA.nodeId, b: pB.nodeId, used: false });
    }

    // ---- 2b. joinTol 以内の近傍ノード間にもエッジを追加（グラフ連結性の補完） ----
    // CLUSTER_TOL より大きいが joinTol 以内のギャップを持つ端点対を繋ぐ。
    // 同一輪郭の別端点を誤接続しないよう、次数1（端点）のノードのみ対象とする。
    {
      // まず現時点での隣接リストを仮構築して次数を調べる
      const tempDegree = new Map();
      for (const edge of edges) {
        tempDegree.set(edge.a, (tempDegree.get(edge.a) || 0) + 1);
        tempDegree.set(edge.b, (tempDegree.get(edge.b) || 0) + 1);
      }
      // 次数1（端点）のノードを抽出してグリッドで近傍検索
      const endpointNodes = nodes.filter(n => (tempDegree.get(n.id) || 0) === 1);
      const cellSize = config.joinTol * 2;
      const epGrid = new Map();
      for (const ep of endpointNodes) {
        const gk = `${Math.floor(ep.x / cellSize)},${Math.floor(ep.y / cellSize)}`;
        if (!epGrid.has(gk)) epGrid.set(gk, []);
        epGrid.get(gk).push(ep.id);
      }
      const bridgeSeen = new Set(edgeSeen);
      for (const ep of endpointNodes) {
        const gx = Math.floor(ep.x / cellSize);
        const gy = Math.floor(ep.y / cellSize);
        let bestId = -1, bestDist = Infinity;
        for (let dx = -1; dx <= 1; dx++) {
          for (let dy = -1; dy <= 1; dy++) {
            const ids = epGrid.get(`${gx + dx},${gy + dy}`) || [];
            for (const id of ids) {
              if (id === ep.id) continue;
              const d = Math.hypot(nodes[id].x - ep.x, nodes[id].y - ep.y);
              if (d <= config.joinTol && d < bestDist) { bestDist = d; bestId = id; }
            }
          }
        }
        if (bestId < 0) continue;
        const key = ep.id < bestId ? `${ep.id}:${bestId}` : `${bestId}:${ep.id}`;
        if (bridgeSeen.has(key)) continue;
        bridgeSeen.add(key);
        edges.push({ id: edges.length, a: ep.id, b: bestId, used: false });
      }
    }

    if (!edges.length) return { polylines: [], matchRows: [] };

    // ---- 3. ノード→エッジの隣接リスト ----
    const nodeToEdges = new Map();
    for (const edge of edges) {
      if (!nodeToEdges.has(edge.a)) nodeToEdges.set(edge.a, []);
      if (!nodeToEdges.has(edge.b)) nodeToEdges.set(edge.b, []);
      nodeToEdges.get(edge.a).push(edge.id);
      nodeToEdges.get(edge.b).push(edge.id);
    }

    // ---- 4. 分岐ノード（次数≠2）でのエッジペアリング最適化 ----
    // 奇数次数ノードや4次以上のノードで、通過ペアを事前に決定する
    const pairedThrough = new Map(); // edgeId → pairedEdgeId (双方向)
    const matchRows = [];

    for (const [nodeId, eids] of nodeToEdges.entries()) {
      const degree = eids.length;
      if (degree === 2) continue; // 正常な通過ノードはスキップ

      const result = bestPairingAllowOdd(eids, nodeId, nodes, edges);

      for (const pair of result.pairs) {
        pairedThrough.set(pair.a, pair.b);
        pairedThrough.set(pair.b, pair.a);
        matchRows.push({
          node: `N${nodeId}`,
          pair: `E${pair.a}↔E${pair.b}`,
          cost: pair.total,
          angle: pair.angle,
          distance: pair.distance,
          gap: pair.gap,
        });
      }
    }

    // ---- 5. エッジ方向ベクトル計算 ----
    function edgeDirectionFromNode(edgeId, nodeId) {
      const edge = edges[edgeId];
      const otherId = otherNode(edge, nodeId);
      const from = nodes[nodeId];
      const to = nodes[otherId];
      const dx = to.x - from.x;
      const dy = to.y - from.y;
      const len = Math.hypot(dx, dy);
      if (len <= config.EPS) return null;
      return { x: dx / len, y: dy / len };
    }

    // ---- 6. 次エッジ選択：ペアリング優先 → 直進角度優先 ----
    function findNextEdge(nodeId, incomingEdgeId, incomingDir) {
      const available = (nodeToEdges.get(nodeId) || []).filter(eid => eid !== incomingEdgeId && !edges[eid].used);
      if (!available.length) return null;

      // ペアリングが確定している場合はそれを優先
      if (incomingEdgeId != null && pairedThrough.has(incomingEdgeId)) {
        const paired = pairedThrough.get(incomingEdgeId);
        if (available.includes(paired) && !edges[paired].used) return paired;
      }

      if (!incomingDir) return available[0];

      // 角度スコアで最もよい（直進に近い）エッジを選択
      let bestId = null;
      let bestScore = Infinity;
      for (const eid of available) {
        const dir = edgeDirectionFromNode(eid, nodeId);
        if (!dir) continue;
        // 反対方向（直進）に近いほどスコアが低い
        const dot = Math.max(-1, Math.min(1, incomingDir.x * dir.x + incomingDir.y * dir.y));
        // dot が -1 に近い（= 180度逆方向 = 直進）ほど良い
        const straightness = 1 + dot; // 0が完全直進、2がUターン
        const lenPenalty = pointDistance(nodes[nodeId], nodes[otherNode(edges[eid], nodeId)]) * 0.002;
        const score = straightness + lenPenalty;
        if (score < bestScore) {
          bestScore = score;
          bestId = eid;
        }
      }
      return bestId;
    }

    // ---- 7. ポリライン追跡 ----
    function traceFrom(startEdgeId, startNodeId) {
      const startPoint = { x: nodes[startNodeId].x, y: nodes[startNodeId].y };
      const pts = [startPoint];
      let currentNodeId = startNodeId;
      let currentEdgeId = startEdgeId;
      let incomingDir = null;
      let guard = 0;

      while (currentEdgeId != null && guard < edges.length * 4) {
        guard += 1;
        const edge = edges[currentEdgeId];
        if (edge.used) break;
        edge.used = true;

        const nextNodeId = otherNode(edge, currentNodeId);
        const from = nodes[currentNodeId];
        const to = nodes[nextNodeId];
        const dx = to.x - from.x;
        const dy = to.y - from.y;
        const len = Math.hypot(dx, dy);
        incomingDir = len <= config.EPS ? incomingDir : { x: dx / len, y: dy / len };

        pts.push({ x: to.x, y: to.y });
        currentNodeId = nextNodeId;

        // 始点に戻ったら閉じる
        if (currentNodeId === startNodeId) break;

        const nextEdgeId = findNextEdge(currentNodeId, currentEdgeId, incomingDir);
        currentEdgeId = nextEdgeId;
      }

      return cleanPolyline(pts);
    }

    // ---- 8. 開始ノード順：端点（次数1）優先、次に分岐ノード（次数≥3） ----
    const startNodeOrder = [...nodeToEdges.keys()].sort((a, b) => {
      const da = (nodeToEdges.get(a) || []).length;
      const db = (nodeToEdges.get(b) || []).length;
      // 次数1（端点）を最優先
      const pa = da === 1 ? 0 : da === 2 ? 2 : 1;
      const pb = db === 1 ? 0 : db === 2 ? 2 : 1;
      if (pa !== pb) return pa - pb;
      return a - b;
    });

    const polylines = [];
    for (const nodeId of startNodeOrder) {
      for (const edgeId of nodeToEdges.get(nodeId) || []) {
        if (edges[edgeId].used) continue;
        const polyline = traceFrom(edgeId, nodeId);
        if (polyline.length >= config.minPolylinePoints) polylines.push(polyline);
      }
    }

    return { polylines, matchRows };
  }


  function pointKey(point, precision = 4) {
    return `${round(point.x, precision)},${round(point.y, precision)}`;
  }

  function dedupeConsecutivePoints(points, tol = config.EPS * 20) {
    const out = [];
    for (const point of points || []) {
      if (!Number.isFinite(point?.x) || !Number.isFinite(point?.y)) continue;
      if (!out.length || pointDistance(out[out.length - 1], point) > tol) {
        out.push({ x: point.x, y: point.y });
      }
    }
    return out;
  }

  function signedTriangleArea(a, b, c) {
    return (b.x - a.x) * (c.y - a.y) - (b.y - a.y) * (c.x - a.x);
  }

  function cleanPolyline(points) {
    const cleaned = dedupeConsecutivePoints(points);
    if (cleaned.length < 2) return cleaned;

    let changed = true;
    while (changed && cleaned.length >= 3) {
      changed = false;
      for (let i = cleaned.length - 2; i >= 1; i--) {
        const prev = cleaned[i - 1];
        const curr = cleaned[i];
        const next = cleaned[i + 1];
        const area = Math.abs(signedTriangleArea(prev, curr, next));
        const span = pointDistance(prev, next);
        const leg = Math.max(pointDistance(prev, curr), pointDistance(curr, next));
        const thin = leg <= config.EPS ? 0 : area / leg;
        // 間引き条件を緩め、ごく短いセグメントかほぼ直線のみ除去
        if (span <= config.minSegmentLength * 0.1 || thin <= config.EPS * 10) {
          cleaned.splice(i, 1);
          changed = true;
        }
      }
    }

    const closes = cleaned.length >= 3 && pointDistance(cleaned[0], cleaned[cleaned.length - 1]) <= config.joinTol;
    const unique = (closes ? cleaned.slice(0, -1) : cleaned.slice()).map(point => ({ x: point.x, y: point.y }));

    if (closes && unique.length >= 3) {
      unique.push({ x: unique[0].x, y: unique[0].y });
    }

    return unique;
  }

  function endpointDirection(points, atStart) {
    if (!points || points.length < 2) return { x: 0, y: 0 };
    if (atStart) {
      for (let i = 1; i < points.length; i++) {
        const dx = points[i].x - points[0].x;
        const dy = points[i].y - points[0].y;
        const len = Math.hypot(dx, dy);
        if (len > config.EPS) return { x: dx / len, y: dy / len };
      }
      return { x: 0, y: 0 };
    }
    for (let i = points.length - 2; i >= 0; i--) {
      const dx = points[points.length - 1].x - points[i].x;
      const dy = points[points.length - 1].y - points[i].y;
      const len = Math.hypot(dx, dy);
      if (len > config.EPS) return { x: dx / len, y: dy / len };
    }
    return { x: 0, y: 0 };
  }

  function directionMismatch(a, b) {
    const la = Math.hypot(a.x, a.y);
    const lb = Math.hypot(b.x, b.y);
    if (la <= config.EPS || lb <= config.EPS) return 0;
    const dot = Math.max(-1, Math.min(1, a.x * b.x + a.y * b.y));
    return 1 - dot;
  }

  function orientForMerge(polyA, polyB, mode) {
    if (mode === "tail-head") return [polyA.slice(), polyB.slice()];
    if (mode === "tail-tail") return [polyA.slice(), polyB.slice().reverse()];
    if (mode === "head-head") return [polyA.slice().reverse(), polyB.slice()];
    return [polyA.slice().reverse(), polyB.slice().reverse()];
  }

  function mergeOrClosePolylinePoints(points, tol) {
    const cleaned = cleanPolyline(points);
    if (cleaned.length >= 3 && pointDistance(cleaned[0], cleaned[cleaned.length - 1]) <= tol) {
      const closed = cleaned.slice();
      closed[closed.length - 1] = { x: closed[0].x, y: closed[0].y };
      return cleanPolyline(closed);
    }
    return cleaned;
  }

  function buildMergeCandidate(a, b, repairTol) {
    const modes = ["tail-head", "tail-tail", "head-head", "head-tail"];
    let best = null;
    // joinTolとrepairTolの最大値を実効許容値とする
    const tol = Math.max(repairTol, config.joinTol);

    for (const mode of modes) {
      const [left, right] = orientForMerge(a.points, b.points, mode);
      if (left.length < 2 || right.length < 2) continue;
      const endA = left[left.length - 1];
      const startB = right[0];
      const dist = pointDistance(endA, startB);
      if (dist > tol) continue;

      const dirA = endpointDirection(left, false);
      const dirB = endpointDirection(right, true);
      const mismatch = directionMismatch(dirA, dirB);
      // 距離と方向ミスマッチの重み付けスコア
      const joinPenalty = dist * 1.0 + mismatch * tol * 1.4;
      const candidate = { mode, dist, mismatch, score: round(joinPenalty, 5), left, right };
      if (!best || candidate.score < best.score || (candidate.score === best.score && candidate.dist < best.dist)) best = candidate;
    }

    return best;
  }

  function mergePolylineObjects(a, b, candidate) {
    const mergedPoints = mergeOrClosePolylinePoints([
      ...candidate.left,
      ...candidate.right.slice(pointDistance(candidate.left[candidate.left.length - 1], candidate.right[0]) <= config.joinTol ? 1 : 0),
    ], Math.max(config.joinTol, config.repairTol));

    return {
      points: mergedPoints,
      repaired: true,
      repairType: "merge",
      repairSteps: [
        ...(a.repairSteps || []),
        ...(b.repairSteps || []),
        `端点結合: ${round(candidate.dist, 4)} mm / 方向差 ${round(candidate.mismatch, 4)}`,
      ],
    };
  }

  function autoRepairPolylines(polylines, repairTol) {
    const effectiveTol = Math.max(repairTol, config.joinTol * 0.85, config.EPS * 20);
    const wideTol = Math.min(effectiveTol * 2.5, config.joinTol * 2.0);

    const items = polylines
      .map(points => ({ points: cleanPolyline(points), repaired: false, repairType: null, repairSteps: [] }))
      .filter(item => item.points.length >= 2);

    // ---- パス1: 各ポリラインの端点が近ければ直接閉じる ----
    for (const item of items) {
      if (!isClosedLike(item.points)) {
        const gap = pointDistance(item.points[0], item.points[item.points.length - 1]);
        if (gap <= effectiveTol) {
          item.points = mergeOrClosePolylinePoints(item.points, effectiveTol);
          item.repaired = true;
          item.repairType = "close";
          item.repairSteps = [`端点閉鎖: ${round(gap, 4)} mm`];
        }
      }
    }

    // ---- パス2・3: 開放輪郭同士のマージ（許容値内のペアを反復結合） ----
    function runMergePass(tol) {
      let changed = true;
      while (changed) {
        changed = false;

        // まず単体で端点が閉じられるか確認
        for (const item of items) {
          if (!isClosedLike(item.points) && item.points.length >= 3) {
            const gap = pointDistance(item.points[0], item.points[item.points.length - 1]);
            if (gap <= tol) {
              item.points = mergeOrClosePolylinePoints(item.points, tol);
              if (!item.repaired) { item.repaired = true; item.repairType = "close"; item.repairSteps = [`端点閉鎖: ${round(gap, 4)} mm`]; }
              else { item.repairType = "close"; item.repairSteps.push(`端点閉鎖: ${round(gap, 4)} mm`); }
              changed = true;
            }
          }
        }

        // 次に2つの開放輪郭同士で最良マージを探す
        let best = null;
        for (let i = 0; i < items.length; i++) {
          if (isClosedLike(items[i].points)) continue;
          for (let j = i + 1; j < items.length; j++) {
            if (isClosedLike(items[j].points)) continue;
            const candidate = buildMergeCandidate(items[i], items[j], tol);
            if (!candidate) continue;
            if (!best || candidate.score < best.candidate.score) best = { i, j, candidate };
          }
        }

        if (best) {
          const merged = mergePolylineObjects(items[best.i], items[best.j], best.candidate);
          // 大きいインデックスを先に削除してからインデックスずれを防ぐ
          items.splice(best.j, 1);
          items[best.i] = merged;
          changed = true;
        }
      }
    }

    runMergePass(effectiveTol);
    if (items.some(item => !isClosedLike(item.points))) runMergePass(wideTol);

    // ---- パス5: 距離無制限の最近傍端点結合 ----
    // 開放輪郭が残っている限り、端点間の最短ペアを繰り返し結合する
    // （STLメッシュの欠損による数十mmのギャップも含め強制閉鎖）
    function runNearestNeighborMerge() {
      for (;;) {
        // 開放輪郭リストを取得（インデックス付き）
        const openIdxs = [];
        for (let i = 0; i < items.length; i++) {
          if (!isClosedLike(items[i].points)) openIdxs.push(i);
        }
        if (openIdxs.length === 0) break;

        // 開放が1件 → 自己端点を強制閉鎖
        if (openIdxs.length === 1) {
          const item = items[openIdxs[0]];
          if (item.points.length >= 3) {
            const gap = pointDistance(item.points[0], item.points[item.points.length - 1]);
            const forceTol = Math.max(wideTol, gap * 1.01);
            item.points = mergeOrClosePolylinePoints(item.points, forceTol);
            if (!item.repaired) { item.repaired = true; item.repairType = "close"; item.repairSteps = [`強制閉鎖: ${round(gap, 4)} mm`]; }
            else { item.repairType = "close"; item.repairSteps.push(`強制閉鎖: ${round(gap, 4)} mm`); }
          }
          break;
        }

        // 全開放端点ペアの中で最短距離のペアを探す
        let bestDist = Infinity;
        let bestIA = -1, bestIB = -1, bestMode = "tail-head";

        for (let ai = 0; ai < openIdxs.length; ai++) {
          const ia = openIdxs[ai];
          const pa = items[ia].points;
          const aStart = pa[0], aEnd = pa[pa.length - 1];
          for (let bi = ai + 1; bi < openIdxs.length; bi++) {
            const ib = openIdxs[bi];
            const pb = items[ib].points;
            const bStart = pb[0], bEnd = pb[pb.length - 1];
            const candidates = [
              { dist: pointDistance(aEnd,   bStart), mode: "tail-head" },
              { dist: pointDistance(aEnd,   bEnd),   mode: "tail-tail" },
              { dist: pointDistance(aStart, bStart), mode: "head-head" },
              { dist: pointDistance(aStart, bEnd),   mode: "head-tail" },
            ];
            for (const c of candidates) {
              if (c.dist < bestDist) { bestDist = c.dist; bestIA = ia; bestIB = ib; bestMode = c.mode; }
            }
          }
        }

        if (bestIA < 0) break;

        const [left, right] = orientForMerge(items[bestIA].points, items[bestIB].points, bestMode);
        const rawMerged = [...left, ...right];
        const mergedItem = {
          points: cleanPolyline(rawMerged),
          repaired: true,
          repairType: "merge",
          repairSteps: [
            ...(items[bestIA].repairSteps || []),
            ...(items[bestIB].repairSteps || []),
            `最近傍結合: ${round(bestDist, 4)} mm`,
          ],
        };

        // マージ後の端点ギャップが縮んでいれば閉鎖
        const afterGap = pointDistance(mergedItem.points[0], mergedItem.points[mergedItem.points.length - 1]);
        if (mergedItem.points.length >= 3 && afterGap <= Math.max(wideTol, afterGap * 0.5 + wideTol)) {
          const closeTol = Math.max(wideTol, afterGap * 1.01);
          mergedItem.points = mergeOrClosePolylinePoints(mergedItem.points, closeTol);
          mergedItem.repairType = "close";
          mergedItem.repairSteps.push(`結合後閉鎖: ${round(afterGap, 4)} mm`);
        }

        // インデックスの大きい方を先に削除（小さい方を置き換え）
        const removeHigh = Math.max(bestIA, bestIB);
        const replaceLow = Math.min(bestIA, bestIB);
        items.splice(removeHigh, 1);
        items[replaceLow] = mergedItem;
      }
    }

    if (items.some(item => !isClosedLike(item.points))) runNearestNeighborMerge();

    return items.map(item => {
      const points = mergeOrClosePolylinePoints(item.points, effectiveTol);
      const closed = isClosedLike(points);
      return {
        ...item,
        points,
        repaired: item.repaired || (closed && !isClosedLike(item.points)),
        repairType: item.repairType || (closed ? "close" : null),
      };
    });
  }

  function summarizeSlicePolylines(polylinesBySlice) {
    let totalClosed = 0, totalOpen = 0;
    const sliceChecks = polylinesBySlice.map(slice => {
      const checks = slice.polylines.map((polyline, idx) => {
        const a = analyzePolyline(polyline);
        if (a.closed) totalClosed += 1; else totalOpen += 1;
        return { polylineIndex: idx + 1, ...a };
      });
      return { sliceIndex: slice.index, z: slice.z, total: checks.length, closedCount: checks.filter(c => c.closed).length, openCount: checks.filter(c => !c.closed).length, checks };
    });
    return { totalClosed, totalOpen, sliceChecks };
  }

  function getPolylinePoints(polyline) { return polyline?.points || []; }
  function rotatePolylines90(polylines) { return polylines; } // 回転廃止（描画正確性優先）
  function boundsForPolylines(polylines) { return boundsOfPoints(polylines.flatMap(getPolylinePoints)); }

  function packSlicesTight(polylinesBySlice, sheetWidth, sheetHeight, margin) {
    // 必ずindex順に処理する
    const ordered = [...polylinesBySlice].sort((a, b) => a.index - b.index);
    const sheets = [];
    let sheetIndex = 1;
    let i = 0;

    while (i < ordered.length) {
      const sheet = { index: sheetIndex, items: [] };
      let curX = margin;
      let curY = margin;
      let rowH = 0;

      // 現在の板にできる限り順番通りに詰める
      while (i < ordered.length) {
        const slice = ordered[i];
        const orig = { rotation: 0, polylines: slice.polylines, bounds: boundsForPolylines(slice.polylines) };
        const candidates = [orig]; // 回転なし固定

        let chosen = null;
        for (const c of candidates) {
          const w = c.bounds.width  + margin * 2;
          const h = c.bounds.height + margin * 2;

          let px = curX, py = curY;
          if (px + w > sheetWidth) {
            px = margin;
            py = curY + rowH + margin;
          }
          if (py + h <= sheetHeight) {
            chosen = { ...c, px, py, w, h };
            break;
          }
        }

        // この板に入らない → 次の板へ
        if (!chosen) break;

        if (chosen.py !== curY) {
          curY = chosen.py;
          rowH = 0;
        }

        const placedPolylines = chosen.polylines.map(p => ({
          ...p,
          placement: { x: chosen.px, y: chosen.py, sourceBounds: chosen.bounds, rotation: chosen.rotation },
        }));

        sheet.items.push({
          sliceIndex: slice.index,
          z: slice.z,
          polylines: placedPolylines,
          sourceBounds: chosen.bounds,
          x: chosen.px,
          y: chosen.py,
          rotation: chosen.rotation,
        });

        curX = chosen.px + chosen.w;
        rowH = Math.max(rowH, chosen.h);
        i++;
      }

      if (!sheet.items.length) break; // 無限ループ防止（部品が大きすぎて入らない場合）
      sheets.push(sheet);
      sheetIndex++;
    }

    return { sheets };
  }

  function polylineToSvgPath(points, dx, dy, sourceBounds) {
    if (!points.length) return "";
    const converted = points.map(p => ({ x: dx + (p.x - sourceBounds.minX), y: dy + (sourceBounds.maxY - p.y) }));
    let d = `M ${round(converted[0].x)} ${round(converted[0].y)}`;
    for (let i = 1; i < converted.length; i++) d += ` L ${round(converted[i].x)} ${round(converted[i].y)}`;
    if (isClosedLike(points)) d += " Z";
    return d;
  }

  function labelPositionForItem(item) {
    return { x: item.x + item.sourceBounds.width / 2, y: item.y + item.sourceBounds.height / 2 };
  }

  // ソート済みスライスリストで「index の次に存在するスライス」を返す
  // sliceMap.get(index + 1) では飛び番があるときに誤った部品を参照してしまう
  function getNextSlice(allSlices, currentIndex) {
    const sorted = [...allSlices].sort((a, b) => a.index - b.index);
    const pos = sorted.findIndex(s => s.index === currentIndex);
    return pos >= 0 && pos + 1 < sorted.length ? sorted[pos + 1] : null;
  }

  function makePreviewSvgString(sheet, sheetWidth, sheetHeight, showLabels, allSlices = [], showNextOutline = true) {
    const paths = [], overlays = [], labels = [];
    for (const item of sheet.items) {
      // 赤：現部品（カット線）
      for (const polyline of item.polylines) {
        const points = getPolylinePoints(polyline);
        const placement = polyline.placement || { x: item.x, y: item.y, sourceBounds: item.sourceBounds };
        const d = polylineToSvgPath(points, placement.x, placement.y, placement.sourceBounds || item.sourceBounds);
        if (!d) continue;
        paths.push(`<path d="${d}" fill="none" stroke="rgb(255,0,0)" stroke-width="0.5" />`);
      }
      // 青：次部品（チェック時のみ）
      if (showNextOutline) {
        const nextSlice = getNextSlice(allSlices, item.sliceIndex);
        if (nextSlice && nextSlice.polylines.length) {
          const nb = boundsOfPoints(nextSlice.polylines.flatMap(getPolylinePoints));
          const dx = item.x + item.sourceBounds.width / 2 - nb.width / 2;
          const dy = item.y + item.sourceBounds.height / 2 - nb.height / 2;
          for (const polyline of nextSlice.polylines) {
            const d = polylineToSvgPath(getPolylinePoints(polyline), dx, dy, nb);
            if (!d) continue;
            overlays.push(`<path d="${d}" fill="none" stroke="rgb(0,0,255)" stroke-width="0.2" />`);
          }
        }
      }
      if (showLabels) {
        const label = labelPositionForItem(item);
        labels.push(`<text x="${round(label.x)}" y="${round(label.y)}" font-size="4.5" text-anchor="middle" dominant-baseline="middle" fill="black">S${String(item.sliceIndex).padStart(2, "0")}</text>`);
      }
    }
    // 板輪郭（黒0.2mm）：プレビューのみ、SVG出力には含まない
    const border = `<rect x="0" y="0" width="${sheetWidth}" height="${sheetHeight}" fill="none" stroke="black" stroke-width="0.2" />`;
    return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${sheetWidth} ${sheetHeight}">${border}<g id="cut">${paths.join("")}</g><g id="overlay">${overlays.join("")}</g><g id="label">${labels.join("")}</g></svg>`;
  }

  function makeCutOnlySvgString(sheet, sheetWidth, sheetHeight, allSlices = []) {
    const paths = [], overlays = [], labels = [];
    for (const item of sheet.items) {

      // ---- 赤：現在の部品（カット線） ----
      for (const polyline of item.polylines) {
        const points = getPolylinePoints(polyline);
        const placement = polyline.placement || { x: item.x, y: item.y, sourceBounds: item.sourceBounds };
        const d = polylineToSvgPath(points, placement.x, placement.y, placement.sourceBounds || item.sourceBounds);
        if (!d) continue;
        paths.push(`<path d="${d}" fill="none" stroke="rgb(255,0,0)" stroke-width="0.01" />`);
      }

      // ---- 青：次番号部品の輪郭（重ね合わせ確認用） ----
      const nextSlice = getNextSlice(allSlices, item.sliceIndex);
      if (nextSlice && nextSlice.polylines.length) {
        const nextAllPoints = nextSlice.polylines.flatMap(getPolylinePoints);
        const nb = boundsOfPoints(nextAllPoints); // 次部品の境界ボックス

        // 現部品と次部品それぞれの中心を一致させるオフセットを計算
        const curCx = item.x + item.sourceBounds.width / 2;
        const curCy = item.y + item.sourceBounds.height / 2;
        // polylineToSvgPath の変換式: x' = dx + (p.x - sb.minX), y' = dy + (sb.maxY - p.y)
        // 次部品の中心 (SVG座標): (nb.width/2, nb.height/2)
        // これを現部品中心 (curCx, curCy) に合わせる dx, dy
        const dx = curCx - nb.width / 2;
        const dy = curCy - nb.height / 2;

        for (const polyline of nextSlice.polylines) {
          const points = getPolylinePoints(polyline);
          const d = polylineToSvgPath(points, dx, dy, nb);
          if (!d) continue;
          overlays.push(`<path d="${d}" fill="none" stroke="rgb(0,0,255)" stroke-width="0.01" />`);
        }
      }

      const label = labelPositionForItem(item);
      labels.push(`<text x="${round(label.x)}" y="${round(label.y)}" font-size="4.5" text-anchor="middle" dominant-baseline="middle" fill="black">S${String(item.sliceIndex).padStart(2, "0")}</text>`);
    }
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${sheetWidth}mm" height="${sheetHeight}mm" viewBox="0 0 ${sheetWidth} ${sheetHeight}"><g id="cut">${paths.join("")}</g><g id="overlay">${overlays.join("")}</g><g id="label">${labels.join("")}</g></svg>`;
  }

  function buildSliceDiagnosticSvg(sliceCheck, sliceData, nextSliceData = null, showNextOutline = true) {
    if (!sliceCheck || !sliceData?.polylines?.length) return "";
    const allPoints = sliceData.polylines.flatMap(getPolylinePoints);
    const bounds = boundsOfPoints(allPoints);
    const pad = 8;
    const width = Math.max(40, bounds.width + pad * 2);
    const height = Math.max(40, bounds.height + pad * 2);
    const parts = [];
    // 赤：現スライス
    sliceData.polylines.forEach((polyline, idx) => {
      const points = getPolylinePoints(polyline), a = analyzePolyline(polyline);
      const converted = points.map(p => ({ x: pad + (p.x - bounds.minX), y: pad + (bounds.maxY - p.y) }));
      if (!converted.length) return;
      let d = `M ${round(converted[0].x)} ${round(converted[0].y)}`;
      for (let i = 1; i < converted.length; i++) d += ` L ${round(converted[i].x)} ${round(converted[i].y)}`;
      if (a.closed) d += " Z";
      parts.push(`<path d="${d}" fill="none" stroke="rgb(255,0,0)" stroke-width="0.5" />`);
      const lp = converted[0];
      parts.push(`<text x="${round(lp.x + 2)}" y="${round(lp.y + 6)}" font-size="4" fill="#475569">P${String(idx + 1).padStart(2, "0")}</text>`);
      if (!a.closed) {
        const start = converted[0], end = converted[converted.length - 1];
        parts.push(`<circle cx="${round(start.x)}" cy="${round(start.y)}" r="2.2" fill="#f59e0b" />`);
        parts.push(`<circle cx="${round(end.x)}" cy="${round(end.y)}" r="2.2" fill="#ef4444" />`);
        parts.push(`<text x="${round((start.x + end.x) / 2 + 3)}" y="${round((start.y + end.y) / 2 - 3)}" font-size="4" fill="#b91c1c">${a.gap ?? "-"}mm</text>`);
      }
    });
    // 青：次スライス（チェック時のみ、中心揃えで重ねる）
    if (showNextOutline && nextSliceData?.polylines?.length) {
      const nb = boundsOfPoints(nextSliceData.polylines.flatMap(getPolylinePoints));
      const cx = pad + bounds.width / 2, cy = pad + bounds.height / 2;
      nextSliceData.polylines.forEach(polyline => {
        const points = getPolylinePoints(polyline);
        const conv = points.map(p => ({
          x: cx - nb.width / 2 + (p.x - nb.minX),
          y: cy - nb.height / 2 + (nb.maxY - p.y),
        }));
        if (!conv.length) return;
        let d = `M ${round(conv[0].x)} ${round(conv[0].y)}`;
        for (let i = 1; i < conv.length; i++) d += ` L ${round(conv[i].x)} ${round(conv[i].y)}`;
        if (isClosedLike(points)) d += " Z";
        parts.push(`<path d="${d}" fill="none" stroke="rgb(0,0,255)" stroke-width="0.2" />`);
      });
    }
    return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${round(width)} ${round(height)}">${parts.join("")}</svg>`;
  }

  function buildSliceDiagnosticSvgBefore(sliceDataBefore, nextSliceData = null, showNextOutline = true) {
    if (!sliceDataBefore?.polylinesBefore?.length) return "";
    const allPoints = sliceDataBefore.polylinesBefore.flatMap(getPolylinePoints);
    const bounds = boundsOfPoints(allPoints);
    const pad = 8;
    const width = Math.max(40, bounds.width + pad * 2);
    const height = Math.max(40, bounds.height + pad * 2);
    const parts = [];
    // 赤：修復前の現スライス
    sliceDataBefore.polylinesBefore.forEach((polyline, idx) => {
      const points = getPolylinePoints(polyline);
      const closed = isClosedLike(points);
      const converted = points.map(p => ({ x: pad + (p.x - bounds.minX), y: pad + (bounds.maxY - p.y) }));
      if (!converted.length) return;
      let d = `M ${round(converted[0].x)} ${round(converted[0].y)}`;
      for (let i = 1; i < converted.length; i++) d += ` L ${round(converted[i].x)} ${round(converted[i].y)}`;
      if (closed) d += " Z";
      parts.push(`<path d="${d}" fill="none" stroke="rgb(255,0,0)" stroke-width="0.5" />`);
      const lp = converted[0];
      parts.push(`<text x="${round(lp.x + 2)}" y="${round(lp.y + 6)}" font-size="4" fill="#475569">P${String(idx + 1).padStart(2, "0")}</text>`);
    });
    // 青：次スライス（チェック時のみ、中心揃えで重ねる）
    if (showNextOutline && nextSliceData?.polylines?.length) {
      const nb = boundsOfPoints(nextSliceData.polylines.flatMap(getPolylinePoints));
      const cx = pad + bounds.width / 2, cy = pad + bounds.height / 2;
      nextSliceData.polylines.forEach(polyline => {
        const points = getPolylinePoints(polyline);
        const conv = points.map(p => ({
          x: cx - nb.width / 2 + (p.x - nb.minX),
          y: cy - nb.height / 2 + (nb.maxY - p.y),
        }));
        if (!conv.length) return;
        let d = `M ${round(conv[0].x)} ${round(conv[0].y)}`;
        for (let i = 1; i < conv.length; i++) d += ` L ${round(conv[i].x)} ${round(conv[i].y)}`;
        if (isClosedLike(points)) d += " Z";
        parts.push(`<path d="${d}" fill="none" stroke="rgb(0,0,255)" stroke-width="0.2" />`);
      });
    }
    return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${round(width)} ${round(height)}">${parts.join("")}</svg>`;
  }

  function buildImprovementProposal(closureSummary) {
    if (!closureSummary?.sliceChecks?.length) return { level: "good", headline: "評価対象の輪郭がありません", items: ["まずSTLを読み込み、スライス生成を実行してください。"] };
    const allChecks = closureSummary.sliceChecks.flatMap(s => s.checks);
    const repairedCount = allChecks.filter(c => c.repaired).length;
    const nearestMergeCount = allChecks.filter(c => c.repairSteps?.some(s => s.includes("最近傍"))).length;
    const worst = [...closureSummary.sliceChecks].sort((a, b) => b.openCount - a.openCount).slice(0, 3).filter(s => s.openCount > 0);
    const items = [];
    let level = "good", headline = "輪郭品質は概ね良好です";
    if (worst.some(s => s.openCount >= 3)) {
      level = "critical";
      headline = "開放輪郭が残っています";
      items.push("パス1–4（許容値内閉鎖・マージ）＋パス5（最近傍グラフ結合）を適用済みです。");
      items.push("残る開放輪郭はSTLメッシュの大きな欠損か、孤立した短い断片です。");
      items.push("接続許容値・修復許容値をさらに上げるか、STLファイル自体を修復してください。");
    }
    items.push(`自動修復済み輪郭: ${repairedCount} 本（うち最近傍結合: ${nearestMergeCount} 本）`);
    if (worst.length) items.push(`重点確認スライス: ${worst.map(s => `S${String(s.sliceIndex).padStart(2, "0")} (開放 ${s.openCount})`).join("、")}`);
    return { level, headline, items };
  }

  function getCutOnlySvg(sheetIndex) {
    const r = state.result;
    return makeCutOnlySvgString(r.sheets[sheetIndex], r.sheetWidth, r.sheetHeight, r.slicesWithPolylines);
  }

  function saveCurrentCutOnlySvg() {
    if (!state.result) return;
    const svg = getCutOnlySvg(state.selectedSheetIndex);
    downloadText(`${state.modelName || "slice"}_sheet_${state.selectedSheetIndex + 1}.svg`, svg);
    setStatus("カット用SVGを保存しました。");
  }

  async function copyCurrentCutOnlySvg() {
    if (!state.result) return;
    const svg = getCutOnlySvg(state.selectedSheetIndex);
    try {
      await navigator.clipboard.writeText(svg);
      setStatus("カット用SVG文字列をクリップボードにコピーしました。");
    } catch {
      setStatus("クリップボードへのコピーに失敗しました。");
    }
  }

  function saveAllCutOnlySvgs() {
    if (!state.result) return;
    state.result.sheets.forEach((_, i) => {
      setTimeout(() => downloadText(`${state.modelName || "slice"}_sheet_${i + 1}.svg`, getCutOnlySvg(i)), i * 180);
    });
    setStatus("全板のSVG保存を開始しました。ブラウザ設定により複数ダウンロードが制限される場合があります。");
  }

  async function saveAllCutOnlyPdf() {
    if (!state.result) return;
    const r = state.result;
    const sheetW = r.sheetWidth;
    const sheetH = r.sheetHeight;

    setStatus("PDF生成中...");
    try {
      // SVGをPDFのXObjectとして直接埋め込む（ベクター品質）
      // PDF座標系: 1mm = 2.8346pt
      const MM2PT = 2.8346;
      const pageW = round(sheetW * MM2PT, 3);
      const pageH = round(sheetH * MM2PT, 3);

      const enc = new TextEncoder();
      const parts = [];
      let offset = 0;
      const offsets = {};

      function wb(str) {
        const b = enc.encode(str);
        parts.push(b); offset += b.length;
      }
      function wBytes(b) {
        parts.push(b); offset += b.length;
      }

      wb("%PDF-1.4\n%\xFF\xFF\n");

      const n = r.sheets.length;
      // オブジェクト割り当て: 1=catalog, 2=pages, 3+i=page, 3+n+i=content
      const catalogId = 1, pagesId = 2;
      const pageIds    = Array.from({length: n}, (_, i) => 3 + i);
      const contentIds = Array.from({length: n}, (_, i) => 3 + n + i);

      // Catalog
      offsets[catalogId] = offset;
      wb(`${catalogId} 0 obj\n<< /Type /Catalog /Pages ${pagesId} 0 R >>\nendobj\n`);

      // Pages
      offsets[pagesId] = offset;
      wb(`${pagesId} 0 obj\n<< /Type /Pages /Kids [`);
      pageIds.forEach(id => wb(`${id} 0 R `));
      wb(`] /Count ${n} >>\nendobj\n`);

      for (let i = 0; i < n; i++) {
        // SVGをPDF content streamにパス変換して埋め込む
        const svgStr = getCutOnlySvg(i);
        const pathCmds = svgToPdfPaths(svgStr, sheetW, sheetH, pageW, pageH, MM2PT);
        const contentBytes = enc.encode(pathCmds);

        // Content stream
        offsets[contentIds[i]] = offset;
        wb(`${contentIds[i]} 0 obj\n<< /Length ${contentBytes.length} >>\nstream\n`);
        wBytes(contentBytes);
        wb(`\nendstream\nendobj\n`);

        // Page object
        offsets[pageIds[i]] = offset;
        wb(`${pageIds[i]} 0 obj\n`);
        wb(`<< /Type /Page /Parent ${pagesId} 0 R `);
        wb(`/MediaBox [0 0 ${pageW} ${pageH}] `);
        wb(`/Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >> `);
        wb(`/Contents ${contentIds[i]} 0 R >>\n`);
        wb(`endobj\n`);
      }

      // xref
      const xrefOffset = offset;
      const totalObjs = 2 + 2 * n;
      wb(`xref\n0 ${totalObjs + 1}\n`);
      wb(`0000000000 65535 f \n`);
      for (let id = 1; id <= totalObjs; id++) {
        wb(`${String(offsets[id] || 0).padStart(10, "0")} 00000 n \n`);
      }
      wb(`trailer\n<< /Size ${totalObjs + 1} /Root ${catalogId} 0 R >>\n`);
      wb(`startxref\n${xrefOffset}\n%%EOF\n`);

      const total = parts.reduce((s, p) => s + p.length, 0);
      const result = new Uint8Array(total);
      let pos = 0;
      for (const p of parts) { result.set(p, pos); pos += p.length; }

      const blob = new Blob([result], { type: "application/pdf" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${state.modelName || "slice"}_all_sheets.pdf`;
      a.style.display = "none";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 2000);
      setStatus(`全板PDFを保存しました（${n}ページ・ベクター品質）`);
    } catch (err) {
      setStatus(`PDF生成に失敗しました: ${err.message || err}`);
    }
  }

  // SVG の path/text要素をPDF座標系の命令に変換
  function svgToPdfPaths(svgStr, sheetW, sheetH, pageW, pageH, mm2pt) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(svgStr, "image/svg+xml");
    const lines = [];

    // PDF座標系はY軸が下から上なのでY反転
    function tx(x) { return round(parseFloat(x) * mm2pt, 4); }
    function ty(y) { return round((sheetH - parseFloat(y)) * mm2pt, 4); }

    // path要素（カット線・輪郭）
    for (const path of doc.querySelectorAll("path")) {
      const stroke = path.getAttribute("stroke") || "rgb(0,0,0)";
      const sw = parseFloat(path.getAttribute("stroke-width") || "0.1");
      const d = path.getAttribute("d") || "";
      const col = parseRgbToPdfColor(stroke);
      lines.push(`${col.r} ${col.g} ${col.b} RG`);
      lines.push(`${round(sw * mm2pt, 4)} w`);
      const cmd = svgPathToPdfOps(d, tx, ty);
      if (cmd) { lines.push(cmd); lines.push("S"); }
    }

    // text要素（部品番号ラベル）
    // PDFにフォントリソースなしでテキストを出すにはType1の組み込みフォント Helvetica を使用
    for (const el of doc.querySelectorAll("text")) {
      const x = parseFloat(el.getAttribute("x") || "0");
      const y = parseFloat(el.getAttribute("y") || "0");
      const fontSize = parseFloat(el.getAttribute("font-size") || "4.5");
      const fill = el.getAttribute("fill") || "black";
      const text = el.textContent || "";
      if (!text.trim()) continue;

      const col = parseRgbToPdfColor(fill === "black" ? "rgb(0,0,0)" : fill);
      const ptSize = round(fontSize * mm2pt, 3);
      const ptX = tx(x);
      const ptY = ty(y);

      // PDFテキスト命令（text-anchor:middleを考慮して左寄せに補正）
      // 文字幅の近似：フォントサイズ × 文字数 × 0.5 で中央補正
      const approxW = ptSize * text.length * 0.5;
      const adjX = round(ptX - approxW / 2, 4);

      lines.push(`${col.r} ${col.g} ${col.b} rg`);   // fill color
      lines.push(`BT`);
      lines.push(`/F1 ${ptSize} Tf`);
      lines.push(`${adjX} ${ptY} Td`);
      // テキストを16進エンコード（ASCIIのみ）
      const hex = Array.from(text).map(c => c.charCodeAt(0).toString(16).padStart(2, "0")).join("");
      lines.push(`<${hex}> Tj`);
      lines.push(`ET`);
    }

    return lines.join("\n") + "\n";
  }

  function parseRgbToPdfColor(str) {
    const m = str.match(/rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/i);
    if (m) return {
      r: round(parseInt(m[1]) / 255, 4),
      g: round(parseInt(m[2]) / 255, 4),
      b: round(parseInt(m[3]) / 255, 4),
    };
    if (str === "black" || str === "#000000") return { r: 0, g: 0, b: 0 };
    return { r: 0, g: 0, b: 0 };
  }

  function svgPathToPdfOps(d, tx, ty) {
    // SVGパスのM/L/Z命令をPDF m/l/h に変換（簡易実装：M/L/Zのみ対応）
    const ops = [];
    const tokens = d.trim().split(/(?=[MLZmlz])|[\s,]+/).filter(Boolean);
    let i = 0;
    while (i < tokens.length) {
      const cmd = tokens[i];
      if (cmd === "M" || cmd === "m") {
        const x = tokens[++i], y = tokens[++i];
        ops.push(`${tx(x)} ${ty(y)} m`);
      } else if (cmd === "L" || cmd === "l") {
        const x = tokens[++i], y = tokens[++i];
        ops.push(`${tx(x)} ${ty(y)} l`);
      } else if (cmd === "Z" || cmd === "z") {
        ops.push("h");
      } else if (!isNaN(parseFloat(cmd))) {
        // 暗黙的なL命令
        const x = cmd, y = tokens[++i];
        ops.push(`${tx(x)} ${ty(y)} l`);
      }
      i++;
    }
    return ops.join("\n");
  }

  function downloadText(filename, text, mime = "image/svg+xml;charset=utf-8") {
    const blob = new Blob([text], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function positiveNumber(value, fallback) {
    const n = Number(value);
    return Number.isFinite(n) && n > 0 ? n : fallback;
  }

  function round(n, d = 4) {
    const p = 10 ** d;
    return Math.round(n * p) / p;
  }

  function pointDistance(a, b) {
    return Math.hypot(a.x - b.x, a.y - b.y);
  }

  function samePoint(a, b, tol = config.joinTol) {
    return pointDistance(a, b) <= tol;
  }

  function isClosedLike(points) {
    return points.length > 2 && samePoint(points[0], points[points.length - 1], config.joinTol);
  }

  function setStatus(text) {
    uiSafe.statusBox.textContent = text;
  }

  function escapeHtml(str) {
    return String(str).replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;").replaceAll("'","&#039;");
  }

  // ---- 3D 組み立て完成図（Canvas 2D 等角投影） ----
  let anim3d = null;
  let rot3dX = 0.5, rot3dY = 0.6, zoom3d = 1.0;
  let is3dDragging = false, last3dMX = 0, last3dMY = 0;
  function start3d() {
    if (anim3d) { cancelAnimationFrame(anim3d); anim3d = null; }
    const canvas = uiSafe.canvas3d;
    const ctx = canvas.getContext("2d");
    const r = state.result;
    if (!r) return;

    // 表示部品数の上限（show3dCount）
    const allSlicesSorted = [...r.slicesWithPolylines].sort((a, b) => a.index - b.index);

    function getVisibleSlices() {
      const count = Math.min(state.show3dCount || allSlicesSorted.length, allSlicesSorted.length);
      return allSlicesSorted.slice(0, count);
    }
    const thickness = positiveNumber(document.getElementById("thicknessInput").value, 3);

    // 全スライスのXY boundsは常に全体で計算（カメラ位置安定のため）
    const allPts = allSlicesSorted.flatMap(s => s.polylines.flatMap(getPolylinePoints));
    const bounds = boundsOfPoints(allPts);
    const modelW = bounds.width, modelD = bounds.height;
    const totalHFull = allSlicesSorted.length * thickness;

    // 初期姿勢リセット
    rot3dX = 0.5; rot3dY = 0.6; zoom3d = 1.0;

    // イベントリスナーは最初の1回だけ登録
    if (!canvas._3dInited) {
      canvas._3dInited = true;
      canvas.addEventListener("mousedown", e => {
        is3dDragging = true; last3dMX = e.clientX; last3dMY = e.clientY;
        canvas.classList.add("grabbing");
      });
      window.addEventListener("mouseup", () => {
        is3dDragging = false; canvas.classList.remove("grabbing");
      });
      window.addEventListener("mousemove", e => {
        if (!is3dDragging || !state.show3d) return;
        rot3dY += (e.clientX - last3dMX) * 0.01;
        rot3dX += (e.clientY - last3dMY) * 0.01;
        rot3dX = Math.max(-Math.PI / 2, Math.min(Math.PI / 2, rot3dX));
        last3dMX = e.clientX; last3dMY = e.clientY;
      });
      canvas.addEventListener("wheel", e => {
        if (!state.show3d) return;
        e.preventDefault();
        zoom3d *= e.deltaY < 0 ? 1.15 : 1 / 1.15;
        zoom3d = Math.max(0.1, Math.min(10, zoom3d));
      }, { passive: false });
    }

    function sliceColor(i, total) {
      const t = i / Math.max(total - 1, 1);
      const rv = Math.round(220 - t * 80);
      const gv = Math.round(80 + t * 100);
      const bv = Math.round(180 - t * 60);
      return `rgb(${rv},${gv},${bv})`;
    }

    function project(x, y, z) {
      const x1 = x * Math.cos(rot3dY) - z * Math.sin(rot3dY);
      const z1 = x * Math.sin(rot3dY) + z * Math.cos(rot3dY);
      const y1 = y * Math.cos(rot3dX) - z1 * Math.sin(rot3dX);
      const z2 = y * Math.sin(rot3dX) + z1 * Math.cos(rot3dX);
      return { sx: x1, sy: -y1, depth: z2 };
    }

    function draw() {
      const W = canvas.offsetWidth, H = canvas.offsetHeight;
      if (canvas.width !== W || canvas.height !== H) { canvas.width = W; canvas.height = H; }
      ctx.clearRect(0, 0, W, H);
      ctx.fillStyle = "#1a1a2e";
      ctx.fillRect(0, 0, W, H);

      // 表示対象スライスを毎フレーム取得（ボタン操作に追従）
      const slices = getVisibleSlices();
      const totalH = slices.length * thickness;
      const yCenter = totalHFull / 2;  // 全体中心を基準にして安定させる

      const maxDim = Math.max(modelW, modelD, totalHFull, 1);
      const baseScale = Math.min(W, H) * 0.55 / maxDim;
      const scale = baseScale * zoom3d;
      const cx = W / 2, cy = H / 2;

      // 底面グリッド
      ctx.strokeStyle = "rgba(255,255,255,0.08)";
      ctx.lineWidth = 0.5;
      const gridStep = Math.max(10, Math.round(maxDim / 5));
      for (let gx = Math.floor(bounds.minX / gridStep) * gridStep; gx <= bounds.maxX; gx += gridStep) {
        const p0 = project(gx - bounds.minX - modelW/2, -yCenter, 0 - modelD/2);
        const p1 = project(gx - bounds.minX - modelW/2, -yCenter, modelD - modelD/2);
        ctx.beginPath();
        ctx.moveTo(cx + p0.sx * scale, cy + p0.sy * scale);
        ctx.lineTo(cx + p1.sx * scale, cy + p1.sy * scale);
        ctx.stroke();
      }

      // 各スライスを板厚分の高さで描画（下から積層）
      slices.forEach((slice, si) => {
        const yBase = si * thickness;
        const yTop  = yBase + thickness;
        const isLast = si === slices.length - 1;  // 最後（一番上）の部品
        const color = sliceColor(si, allSlicesSorted.length);
        const colorDark = sliceColor(si, allSlicesSorted.length).replace("rgb(", "rgba(").replace(")", ",0.5)");

        slice.polylines.forEach(poly => {
          const pts = getPolylinePoints(poly);
          if (pts.length < 2) return;

          // 上面の投影点を計算
          const topProj = pts.map(p => {
            const pr = project(p.x - bounds.minX - modelW/2, yTop - yCenter, p.y - bounds.minY - modelD/2);
            return { x: cx + pr.sx * scale, y: cy + pr.sy * scale };
          });

          // 最後の部品は上面を黄色で塗りつぶし
          if (isLast && isClosedLike(pts) && topProj.length >= 3) {
            ctx.beginPath();
            ctx.moveTo(topProj[0].x, topProj[0].y);
            for (let i = 1; i < topProj.length; i++) ctx.lineTo(topProj[i].x, topProj[i].y);
            ctx.closePath();
            ctx.fillStyle = "rgba(255, 220, 0, 0.75)";
            ctx.fill();
          }

          // 輪郭線
          ctx.beginPath();
          ctx.moveTo(topProj[0].x, topProj[0].y);
          for (let i = 1; i < topProj.length; i++) ctx.lineTo(topProj[i].x, topProj[i].y);
          if (isClosedLike(pts)) ctx.closePath();
          ctx.strokeStyle = isLast ? "#ffdd00" : color;
          ctx.lineWidth = isLast ? 1.5 : 1;
          ctx.stroke();

          // 側面（板厚の垂直線）
          const step = Math.max(1, Math.floor(pts.length / 24));
          for (let i = 0; i < pts.length; i += step) {
            const p = pts[i];
            const pB = project(p.x - bounds.minX - modelW/2, yBase - yCenter, p.y - bounds.minY - modelD/2);
            const pT = project(p.x - bounds.minX - modelW/2, yTop  - yCenter, p.y - bounds.minY - modelD/2);
            ctx.beginPath();
            ctx.moveTo(cx + pT.sx * scale, cy + pT.sy * scale);
            ctx.lineTo(cx + pB.sx * scale, cy + pB.sy * scale);
            ctx.strokeStyle = isLast ? "rgba(255,220,0,0.5)" : colorDark;
            ctx.lineWidth = 0.6;
            ctx.stroke();
          }
        });
      });

      // ヘルプテキスト
      ctx.fillStyle = "rgba(255,255,255,0.35)";
      ctx.font = "12px system-ui";
      const countLabel = slices.length >= allSlicesSorted.length ? `全${allSlicesSorted.length}部品` : `${slices.length}/${allSlicesSorted.length}部品`;
      ctx.fillText(`ドラッグで回転 / ホイールで拡大縮小  ${Math.round(zoom3d * 100)}%  ${countLabel}`, 12, H - 12);

      if (state.show3d) anim3d = requestAnimationFrame(draw);
    }

    draw();
  }

  if (typeof window !== "undefined") {
    window.__sliceDebug = {
      tracePolylinesV7,
      autoRepairPolylines,
      cleanPolyline,
      buildSegments,
      parseAsciiSTL,
      parseBinarySTL,
      setConfig: values => Object.assign(config, values || {}),
    };
  }
}
