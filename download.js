(() => {
  'use strict';

  const FILE_URL = 'files/sports-culture-2026.pptx';
  const SAVE_NAME = '体育文化传承与互鉴_2026-11-13.pptx';
  const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
  const byId = id => document.getElementById(id);
  const ui = {
    button: byId('downloadButton'), label: byId('buttonLabel'),
    status: byId('downloadStatus'), wrap: byId('progressWrap'),
    progress: byId('downloadProgress'), detail: byId('progressDetail'),
    cancel: byId('cancelButton')
  };
  let phase = 'idle';
  let controller = null;
  let readyBlob = null;
  let manifest = null;
  let manifestPromise;
  let retryNeedsMetadata = false;
  let preferBrowserSave = false;

  function formatBytes(bytes) {
    if (bytes < 1024) return `${bytes.toLocaleString('zh-CN')} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  }

  function setStatus(message, tone = '') {
    ui.status.textContent = message;
    ui.status.dataset.tone = tone;
  }

  function setButton(label, disabled = false) {
    ui.label.textContent = label;
    ui.button.disabled = disabled;
  }

  function showProgress(loaded, total) {
    ui.wrap.hidden = false;
    if (total !== null && loaded <= total) {
      ui.progress.max = total;
      ui.progress.value = loaded;
      ui.detail.textContent = `${formatBytes(loaded)} / ${formatBytes(total)} · ${Math.floor(loaded / total * 100)}%`;
    } else {
      ui.progress.removeAttribute('value');
      ui.detail.textContent = loaded > 0 ? `已接收 ${formatBytes(loaded)} · 总大小未知` : '正在连接…';
    }
  }

  async function loadManifest() {
    const metadataController = new AbortController();
    const metadataTimeout = setTimeout(() => metadataController.abort(), 8000);
    try {
      const response = await fetch('manifest.json', { cache: 'no-cache', signal: metadataController.signal });
      if (!response.ok) return;
      const data = await response.json();
      if (!data || typeof data !== 'object') return;
      manifest = data;
      if (Number.isInteger(data.slides) && data.slides > 0) byId('slideCount').textContent = `${data.slides} 页`;
      if (Number.isFinite(data.bytes) && data.bytes > 0) byId('fileSize').textContent = formatBytes(data.bytes);
      if (typeof data.duration === 'string') byId('duration').textContent = data.duration;
      if (typeof data.compatibility === 'string') byId('compatibility').textContent = data.compatibility;
      if (typeof data.sha256 === 'string' && /^[a-f\d]{64}$/i.test(data.sha256)) {
        byId('checksumDetails').hidden = false;
        byId('checksumValue').textContent = data.sha256.toLowerCase();
      }
      if (data.preview === true) byId('previewLink').hidden = false;
      if (data.notes === true) byId('notesLink').hidden = false;
      byId('secondaryFiles').hidden = data.preview !== true && data.notes !== true;
    } catch {
      // File retrieval remains usable when optional metadata is unavailable.
    } finally {
      clearTimeout(metadataTimeout);
    }
  }

  async function validateFile(blob, received, transferTotal) {
    if (received === 0) throw new Error('服务器返回了空文件。');
    if (transferTotal !== null && received !== transferTotal) throw new Error('接收字节与服务器声明的大小不一致。');
    await manifestPromise;
    if (Number.isSafeInteger(manifest?.bytes) && manifest.bytes > 0 && manifest.bytes !== received) {
      throw new Error('接收文件的大小与发布记录不一致。');
    }
    const head = new Uint8Array(await blob.slice(0, 4).arrayBuffer());
    if (head.length !== 4 || head[0] !== 0x50 || head[1] !== 0x4b || head[2] !== 0x03 || head[3] !== 0x04) {
      throw new Error('返回内容不是有效的 PPTX 文件。');
    }
    if (/^[a-f\d]{64}$/i.test(manifest?.sha256 ?? '') && globalThis.crypto?.subtle) {
      const digest = await crypto.subtle.digest('SHA-256', await blob.arrayBuffer());
      const hash = Array.from(new Uint8Array(digest), byte => byte.toString(16).padStart(2, '0')).join('');
      if (hash !== manifest.sha256.toLowerCase()) throw new Error('文件校验失败，请重新下载。');
    }
  }

  async function prepareDownload() {
    phase = 'loading';
    if (retryNeedsMetadata) {
      manifestPromise = loadManifest();
      retryNeedsMetadata = false;
    }
    readyBlob = null;
    controller = new AbortController();
    const requestController = controller;
    setButton('正在读取演示文稿…', true);
    setStatus('正在读取文件；尚未保存到设备。');
    ui.cancel.hidden = false;
    showProgress(0, null);
    try {
      const response = await fetch(FILE_URL, { signal: requestController.signal, cache: 'no-cache' });
      if (!response.ok) throw new Error(`服务器返回 HTTP ${response.status}。`);
      const length = response.headers.get('content-length');
      const encoding = response.headers.get('content-encoding');
      const parsedLength = length && /^\d+$/.test(length) ? Number(length) : NaN;
      const total = Number.isSafeInteger(parsedLength) && parsedLength > 0 && (!encoding || encoding === 'identity') ? parsedLength : null;
      let received = 0;
      let blob;
      if (response.body && typeof response.body.getReader === 'function') {
        const reader = response.body.getReader();
        const chunks = [];
        while (true) {
          const { done, value } = await reader.read();
          if (done) break;
          chunks.push(value);
          received += value.byteLength;
          showProgress(received, total);
        }
        blob = new Blob(chunks, { type: PPTX_MIME });
      } else {
        // Older browsers cannot expose streamed bytes; never invent progress.
        ui.detail.textContent = '浏览器正在读取文件 · 暂无法显示分段进度';
        blob = await response.blob();
        received = blob.size;
      }
      ui.cancel.hidden = true;
      setButton('正在检查文件…', true);
      setStatus('已接收文件，正在检查完整性；尚未保存到设备。');
      await validateFile(blob, received, total);
      if (requestController.signal.aborted) throw new DOMException('Aborted', 'AbortError');
      readyBlob = blob;
      phase = 'ready';
      ui.wrap.hidden = true;
      setButton('保存演示文稿');
      setStatus(`已接收 ${formatBytes(received)}，已准备保存。点击上方按钮选择保存。`, 'success');
    } catch (error) {
      readyBlob = null;
      phase = 'idle';
      retryNeedsMetadata = true;
      ui.wrap.hidden = true;
      if (error.name === 'AbortError') {
        setButton('重新下载演示文稿');
        setStatus('已取消读取。文件尚未准备完成，也未保存到设备。');
      } else {
        setButton('重试下载');
        const detail = error instanceof TypeError ? '网络连接中断或请求无法完成。' : error.message;
        setStatus(`下载未完成：${detail} 请重试，或使用“浏览器直接下载”。`, 'error');
      }
    } finally {
      controller = null;
    }
  }

  async function savePreparedFile() {
    if (!readyBlob) return;
    phase = 'saving';
    setButton('正在保存…', true);
    if (preferBrowserSave) {
      // A fresh user click invokes the same visible native link that works in restricted browsers.
      try {
        byId('nativeLink').click();
        setStatus('已请求浏览器直接下载，请在下载列表确认。网页无法确认文件已写入设备。');
        setButton('再次使用浏览器下载');
      } catch {
        setStatus('未能启动浏览器下载。请点击下方“浏览器直接下载”链接。', 'error');
        setButton('使用浏览器下载');
      }
      phase = 'ready';
      return;
    }
    if (window.isSecureContext && typeof window.showSaveFilePicker === 'function') {
      let writable;
      try {
        // Invoked directly by a fresh click: saving always remains the visitor's choice.
        const handle = await window.showSaveFilePicker({
          suggestedName: SAVE_NAME,
          types: [{ description: 'PowerPoint 演示文稿', accept: { [PPTX_MIME]: ['.pptx'] } }]
        });
        writable = await handle.createWritable();
        await writable.write(readyBlob);
        await writable.close();
        writable = null;
        setStatus('已写入你选择的文件。可在所选位置打开演示文稿。', 'success');
        setButton('另存一份');
      } catch (error) {
        if (writable) { try { await writable.abort(); } catch { /* Preserve the original error. */ } }
        if (error.name === 'AbortError') {
          setStatus('已取消保存。文件仍已准备好，可再次点击保存。');
          setButton('保存演示文稿');
        } else {
          // Do not repeat a picker that exists but cannot be used in this environment.
          // Wait for another explicit click; cancellation never starts a fallback download.
          preferBrowserSave = true;
          const unavailable = ['SecurityError', 'NotAllowedError', 'NotSupportedError'].includes(error.name);
          setStatus(`${unavailable ? '当前环境无法使用系统保存。' : '保存未完成。'}文件仍已准备好，可点击“使用浏览器下载”，或使用下方直接下载链接。`, 'error');
          setButton('使用浏览器下载');
        }
      } finally {
        phase = 'ready';
      }
    } else {
      try {
        const objectUrl = URL.createObjectURL(readyBlob);
        const link = document.createElement('a');
        link.href = objectUrl;
        link.download = SAVE_NAME;
        link.hidden = true;
        document.body.appendChild(link);
        link.click();
        link.remove();
        setTimeout(() => URL.revokeObjectURL(objectUrl), 60000);
        setStatus('已交给浏览器保存，请在下载列表确认。网页无法确认文件已写入设备。');
        setButton('再次保存');
      } catch {
        setStatus('未能启动浏览器保存。请使用“浏览器直接下载”。', 'error');
        setButton('保存演示文稿');
      }
      phase = 'ready';
    }
  }

  ui.button.addEventListener('click', () => {
    if (phase === 'idle') return prepareDownload();
    if (phase === 'ready') return savePreparedFile();
  });
  ui.cancel.addEventListener('click', () => controller?.abort());
  byId('nativeLink').addEventListener('click', () => {
    setStatus('已请求浏览器直接下载。请在浏览器下载列表查看进度和保存结果。');
  });
  manifestPromise = loadManifest();
})();
