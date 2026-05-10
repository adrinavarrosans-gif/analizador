"""
gcs_uploader_component.py — Custom Streamlit HTML component for direct
browser-to-GCS uploads via signed URLs.

Usage in app.py:
    from gcs_uploader_component import render_gcs_uploader
    render_gcs_uploader(signed_url, blob_name, label, key)
"""

import streamlit as st
import streamlit.components.v1 as components


def render_gcs_uploader(signed_url: str, blob_name: str,
                        label: str = "Selecciona archivo PDF",
                        key: str = "gcs_upload",
                        accept: str = ".pdf",
                        max_size_mb: int = 2048) -> None:
    """
    Render a custom HTML/JS file uploader that sends the file directly
    to Google Cloud Storage via a pre-signed PUT URL.

    After a successful upload the JS writes the blob name into
    sessionStorage so that Streamlit can detect completion on the
    next rerun (triggered by the user clicking "Procesar").
    """

    session_key = f"gcs_done_{key}"

    html = f"""
    <div id="gcs-uploader-{key}" style="
        border: 2px dashed #44BABC;
        border-radius: 12px;
        padding: 24px 18px;
        text-align: center;
        background: linear-gradient(135deg, #f0fafa 0%, #ffffff 100%);
        font-family: 'Source Sans Pro', sans-serif;
        min-height: 120px;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        gap: 10px;
        transition: border-color 0.3s;
    ">
        <!-- Estado: seleccionar archivo -->
        <div id="pick-{key}">
            <div style="font-size: 2rem; margin-bottom: 6px;">📂</div>
            <label for="file-{key}" style="
                display: inline-block;
                background: linear-gradient(135deg, #0D5F5D, #44BABC);
                color: white;
                padding: 10px 28px;
                border-radius: 8px;
                cursor: pointer;
                font-weight: 600;
                font-size: 0.95rem;
                box-shadow: 0 2px 8px rgba(13,95,93,0.3);
                transition: transform 0.15s;
            ">{label}</label>
            <input type="file" id="file-{key}" accept="{accept}"
                   style="display:none"
                   onchange="handleFile_{key.replace('-','_')}(this)" />
            <div style="margin-top: 6px; font-size: 0.8rem; color: #888;">
                Máximo {max_size_mb} MB · PDF
            </div>
        </div>

        <!-- Estado: subiendo -->
        <div id="uploading-{key}" style="display:none; width: 100%;">
            <div style="font-size: 1.5rem; margin-bottom: 8px;">☁️</div>
            <div id="filename-{key}" style="font-weight: 600; color: #0D5F5D; margin-bottom: 8px;
                 max-width: 100%; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;"></div>
            <div style="width: 100%; background: #e0e0e0; border-radius: 8px; height: 22px; overflow: hidden;">
                <div id="bar-{key}" style="
                    width: 0%;
                    height: 100%;
                    background: linear-gradient(90deg, #0D5F5D, #44BABC);
                    border-radius: 8px;
                    transition: width 0.3s;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    color: white;
                    font-size: 0.75rem;
                    font-weight: 700;
                "></div>
            </div>
            <div id="status-{key}" style="margin-top: 8px; font-size: 0.85rem; color: #555;"></div>
        </div>

        <!-- Estado: completado -->
        <div id="done-{key}" style="display:none;">
            <div style="font-size: 2.2rem;">✅</div>
            <div style="font-weight: 700; color: #0D5F5D; font-size: 1rem; margin-top: 4px;">
                Archivo subido correctamente
            </div>
            <div id="done-name-{key}" style="font-size: 0.85rem; color: #555; margin-top: 4px;"></div>
            <button onclick="resetUploader_{key.replace('-','_')}()" style="
                margin-top: 10px;
                background: none;
                border: 1px solid #ccc;
                border-radius: 6px;
                padding: 5px 16px;
                cursor: pointer;
                font-size: 0.8rem;
                color: #666;
            ">🔄 Cambiar archivo</button>
        </div>

        <!-- Estado: error -->
        <div id="error-{key}" style="display:none;">
            <div style="font-size: 2rem;">❌</div>
            <div id="error-msg-{key}" style="color: #c0392b; font-weight: 600; font-size: 0.9rem; margin-top: 4px;"></div>
            <button onclick="resetUploader_{key.replace('-','_')}()" style="
                margin-top: 10px;
                background: linear-gradient(135deg, #0D5F5D, #44BABC);
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 20px;
                cursor: pointer;
                font-weight: 600;
                font-size: 0.85rem;
            ">Reintentar</button>
        </div>
    </div>

    <script>
    (function() {{
        const KEY = "{key}";
        const BLOB = "{blob_name}";
        const SIGNED_URL = "{signed_url}";
        const MAX_MB = {max_size_mb};
        const SAFE_KEY = KEY.replace(/-/g, '_');

        function $(id) {{ return document.getElementById(id); }}

        function showState(state) {{
            ['pick', 'uploading', 'done', 'error'].forEach(s => {{
                const el = $(s + '-' + KEY);
                if (el) el.style.display = (s === state) ? '' : 'none';
            }});
        }}

        // Expose as global function for onclick
        window['handleFile_' + SAFE_KEY] = function(input) {{
            const file = input.files[0];
            if (!file) return;

            // Size check
            const sizeMB = file.size / (1024 * 1024);
            if (sizeMB > MAX_MB) {{
                $('error-msg-' + KEY).textContent =
                    'Archivo demasiado grande: ' + sizeMB.toFixed(1) + ' MB (máx: ' + MAX_MB + ' MB)';
                showState('error');
                return;
            }}

            $('filename-' + KEY).textContent = file.name + ' (' + sizeMB.toFixed(1) + ' MB)';
            showState('uploading');

            const xhr = new XMLHttpRequest();
            xhr.open('PUT', SIGNED_URL, true);
            xhr.setRequestHeader('Content-Type', 'application/pdf');

            xhr.upload.onprogress = function(e) {{
                if (e.lengthComputable) {{
                    const pct = Math.round((e.loaded / e.total) * 100);
                    const bar = $('bar-' + KEY);
                    bar.style.width = pct + '%';
                    bar.textContent = pct + '%';
                    const loaded = (e.loaded / (1024*1024)).toFixed(1);
                    const total = (e.total / (1024*1024)).toFixed(1);
                    $('status-' + KEY).textContent =
                        'Subiendo directamente a Cloud Storage: ' + loaded + ' / ' + total + ' MB';
                }}
            }};

            xhr.onload = function() {{
                if (xhr.status >= 200 && xhr.status < 300) {{
                    // Success! Store blob name so Streamlit can find it
                    sessionStorage.setItem('gcs_blob_' + KEY, BLOB);
                    sessionStorage.setItem('gcs_filename_' + KEY, file.name);
                    sessionStorage.setItem('gcs_size_' + KEY, String(file.size));
                    $('done-name-' + KEY).textContent = file.name + ' (' + sizeMB.toFixed(1) + ' MB)';
                    showState('done');

                    // Also store a flag that parent Streamlit can detect
                    try {{
                        window.parent.postMessage({{
                            type: 'gcs_upload_complete',
                            key: KEY,
                            blob: BLOB,
                            filename: file.name,
                            size: file.size
                        }}, '*');
                    }} catch(e) {{}}
                }} else {{
                    $('error-msg-' + KEY).textContent =
                        'Error al subir (HTTP ' + xhr.status + '). Reintenta.';
                    showState('error');
                }}
            }};

            xhr.onerror = function() {{
                $('error-msg-' + KEY).textContent =
                    'Error de conexión. Verifica tu red e inténtalo de nuevo.';
                showState('error');
            }};

            xhr.send(file);
        }};

        window['resetUploader_' + SAFE_KEY] = function() {{
            sessionStorage.removeItem('gcs_blob_' + KEY);
            sessionStorage.removeItem('gcs_filename_' + KEY);
            sessionStorage.removeItem('gcs_size_' + KEY);
            const input = $('file-' + KEY);
            if (input) input.value = '';
            showState('pick');
        }};

        // On load, check if already uploaded
        const existing = sessionStorage.getItem('gcs_blob_' + KEY);
        if (existing === BLOB) {{
            const fname = sessionStorage.getItem('gcs_filename_' + KEY) || '';
            const fsize = parseInt(sessionStorage.getItem('gcs_size_' + KEY) || '0');
            const sizeMB = (fsize / (1024*1024)).toFixed(1);
            $('done-name-' + KEY).textContent = fname + (fsize > 0 ? ' (' + sizeMB + ' MB)' : '');
            showState('done');
        }}
    }})();
    </script>
    """
    components.html(html, height=200, scrolling=False)


def render_gcs_uploader_status_check(key: str, blob_name: str) -> bool:
    """
    Check if the GCS upload for the given key is complete by querying
    GCS directly. Returns True if the blob exists and has content.

    Call this after the user clicks "Procesar".
    """
    from gcs_upload import blob_exists, get_blob_size
    if not blob_name:
        return False
    exists = blob_exists(blob_name)
    if exists:
        size = get_blob_size(blob_name)
        return size > 0
    return False
