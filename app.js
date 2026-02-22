let productosBase = [];
const { jsPDF } = window.jspdf;

// Asegurar que el objeto de resumen exista siempre para evitar errores de descarga
window.resumenInventario = { fisico: 0, sistema: 0, difPesos: 0, porcentaje: "0.00" };

window.onload = () => {
    ['bodega', 'site', 'fecha', 'responsable'].forEach(id => {
        const el = document.getElementById(id);
        if(el) el.value = localStorage.getItem(`meta-${id}`) || "";
    });
};

function guardarMeta() {
    ['bodega', 'site', 'fecha', 'responsable'].forEach(id => {
        const el = document.getElementById(id);
        if(el) localStorage.setItem(`meta-${id}`, el.value);
    });
}

function formatearFechaChile(fechaISO) {
    if (!fechaISO) return "S/F";
    const partes = fechaISO.split('-');
    return partes.length === 3 ? `${partes[2]} - ${partes[1]} - ${partes[0]}` : fechaISO;
}

function leerArchivo(input) {
    const file = input.files[0];
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const filas = XLSX.utils.sheet_to_json(worksheet, {header: 1});

            productosBase = filas.slice(1).map(col => {
                if (!col[0]) return null;
                return {
                    codigo: String(col[0] || "").trim(),
                    nombre: String(col[1] || "").trim(),
                    lote: String(col[2] || "").trim(),
                    un: String(col[4] || "").trim(),
                    teorico: parseFloat(col[5]) || 0,
                    precio: Math.round(parseFloat(col[6]) || 0)
                };
            }).filter(p => p !== null);
            renderizarTarjetas(productosBase);
        } catch (error) {
            alert("Error al leer el Excel.");
        }
    };
    reader.readAsArrayBuffer(file);
}

function renderizarTarjetas(lista) {
    const container = document.getElementById('product-list');
    container.innerHTML = lista.map(p => {
        const key = `inv-${p.codigo}-${p.lote}`;
        const val = localStorage.getItem(key);
        const fisicoVal = val !== null ? parseFloat(val) : 0;
        const inputVal = val !== null ? val : "";
        
        const dif = fisicoVal - p.teorico;
        const colorDif = dif < 0 ? "txt-rojo" : (dif > 0 ? "txt-verde" : "txt-neutral");

        return `
        <div class="product-card">
            <div class="header-card"><span>ID: ${p.codigo}</span><span class="lote-val">${p.lote}</span><span>${p.un}</span></div>
            <div class="product-name">${p.nombre}</div>
            <div class="audit-grid">
                <div class="audit-item"><label>Sist.</label><span>${p.teorico}</span></div>
                <div class="audit-item"><label>Físico</label>
                    <input type="number" inputmode="decimal" value="${inputVal}" placeholder="0" 
                           oninput="actualizarConteo(this, ${p.teorico}, ${p.precio}, '${key}')">
                </div>
                <div class="audit-item"><label>Total Fís.</label><span class="v-total">${formatearMoneda(fisicoVal * p.precio)}</span></div>
            </div>
            <div class="dif-container">
                VALOR DIFERENCIA: <span class="val-dif-pesos ${colorDif}">${formatearMoneda(dif * p.precio)}</span>
                <div style="font-size:0.75em;">CANT. DIF: <span class="cant-dif-val ${colorDif}">${dif.toFixed(2)}</span></div>
            </div>
        </div>`;
    }).join('');
    actualizarTotalesGenerales();
    actualizarBarraProgreso(); // <--- AÑADE ESTO AQUÍ TAMBIÉN
}

function actualizarConteo(input, teorico, precio, key) {
    if (input.value === "") localStorage.removeItem(key);
    else localStorage.setItem(key, input.value);
    
    const fisicoVal = parseFloat(input.value) || 0;
    const dif = fisicoVal - teorico;
    const card = input.closest('.product-card');
    
    card.querySelector('.v-total').innerText = formatearMoneda(fisicoVal * precio);
    const valD = card.querySelector('.val-dif-pesos');
    const cantD = card.querySelector('.cant-dif-val');
    
    valD.innerText = formatearMoneda(dif * precio);
    cantD.innerText = dif.toFixed(2);

    const color = dif < 0 ? "txt-rojo" : (dif > 0 ? "txt-verde" : "txt-neutral");
    valD.className = "val-dif-pesos " + color;
    cantD.className = "cant-dif-val " + color;
    
    actualizarTotalesGenerales();
    actualizarBarraProgreso(); // <--- AGREGA ESTA LÍNEA AQUÍ
}

function actualizarTotalesGenerales() {
    let tSist = 0, tFis = 0;
    productosBase.forEach(p => {
        const val = localStorage.getItem(`inv-${p.codigo}-${p.lote}`);
        const f = val !== null ? parseFloat(val) : 0;
        tSist += (p.teorico * p.precio);
        tFis += (f * p.precio);
    });

    const difP = tFis - tSist;
    const porc = tSist !== 0 ? (Math.abs(difP) / tSist) * 100 : 0;

    // Guardar en objeto global para las descargas
    window.resumenInventario = { fisico: tFis, sistema: tSist, difPesos: difP, porcentaje: porc.toFixed(2) };

    // Actualizar pantalla
    const granTotal = document.getElementById('gran-total');
    if(granTotal) granTotal.innerText = formatearMoneda(tFis);

    const infoExtra = document.getElementById('info-extra-pantalla');
    if(infoExtra) {
        const color = difP < 0 ? 'txt-rojo' : (difP > 0 ? 'txt-verde' : '');
        infoExtra.innerHTML = `Dif: <span class="${color}">${formatearMoneda(difP)}</span> | Ajuste: <span class="${color}">${porc.toFixed(2)}%</span>`;
    }
}

function exportarExcel() {
    try {
        const res = window.resumenInventario;
        const meta = { 
            b: document.getElementById('bodega').value || "S/N", 
            s: document.getElementById('site').value || "S/S", 
            f: formatearFechaChile(document.getElementById('fecha').value),
            r: document.getElementById('responsable').value || "S/R" 
        };

        const fM_Cabecera = (n) => new Intl.NumberFormat('es-CL', { 
            style: 'currency', currency: 'CLP', maximumFractionDigits: 0 
        }).format(n);

        const data = [
            ["REPORTE DE INVENTARIO PROFESIONAL"],
            [`Bodega: ${meta.b}`, `Site: ${meta.s}`, `Fecha: ${meta.f}`, `Responsable: ${meta.r}`],
            [`Total Sistema:`, fM_Cabecera(res.sistema), `Total Físico:`, fM_Cabecera(res.fisico), `% Ajuste:`, parseFloat(res.porcentaje) / 100],
            [],
            ["Código", "Producto", "UN", "Sist.", "Físico", "Dif. Cant", "Vr. Unitario", "Vr. Diferencia", "Total Físico"]
        ];

        productosBase.forEach(p => {
            const f = parseFloat(localStorage.getItem(`inv-${p.codigo}-${p.lote}`)) || 0;
            const d = f - p.teorico;
            data.push([p.codigo, p.nombre, p.un, p.teorico, f, d, p.precio, (d * p.precio), (f * p.precio)]);
        });

        const ws = XLSX.utils.aoa_to_sheet(data);
        const wb = XLSX.utils.book_new();

        // --- FORMATOS CON COLORES DINÁMICOS ---
        // Formato: [Color Positivo];[Color Negativo];[Color Cero]
        const fmtMonedaColor = '[Color10]"$"#,##0.00;[Red]"-" "$"#,##0.00;[Black]"$"0.00'; 
        const fmtCantColor = '[Color10]0.00;[Red]-0.00;[Black]0.00';
        const fmtPorcentaje = '[Color10]0.00%;[Red]-0.00%;[Black]0.00%';

        const range = XLSX.utils.decode_range(ws['!ref']);
        
        // 1. Color al % de Ajuste en cabecera
        const cellPorc = ws[XLSX.utils.encode_cell({r: 2, c: 5})];
        if (cellPorc) cellPorc.z = fmtPorcentaje;

        // 2. Colores en la tabla
        for (let R = 4; R <= range.e.r; ++R) {
            // Columna Dif Cant (F - índice 5)
            const cDif = ws[XLSX.utils.encode_cell({r: R, c: 5})];
            if (cDif) cDif.z = fmtCantColor;

            // Columna Vr. Unitario (G - índice 6) -> Solo moneda normal
            const cUnit = ws[XLSX.utils.encode_cell({r: R, c: 6})];
            if (cUnit) cUnit.z = '"$"#,##0.00';

            // Columnas Vr. Diferencia (H) y Total Físico (I) -> Con colores
            for (let C = 7; C <= 8; ++C) {
                const cell = ws[XLSX.utils.encode_cell({r: R, c: C})];
                if (cell && cell.t === 'n') cell.z = fmtMonedaColor;
            }
        }

        ws['!cols'] = [{wch: 12}, {wch: 35}, {wch: 6}, {wch: 10}, {wch: 10}, {wch: 10}, {wch: 15}, {wch: 15}, {wch: 15}];
        XLSX.utils.book_append_sheet(wb, ws, "Inventario");
        XLSX.writeFile(wb, `Reporte_Excel_${meta.b}.xlsx`);
    } catch (e) { 
        console.error(e);
        alert("Error al generar Excel."); 
    }
}

function exportarPDF() {
    try {
        const doc = new jsPDF('l', 'mm', 'a4');
        const res = window.resumenInventario;
        const meta = { 
            b: document.getElementById('bodega').value || "S/N", 
            s: document.getElementById('site').value || "S/S", 
            f: formatearFechaChile(document.getElementById('fecha').value),
            r: document.getElementById('responsable').value || "S/R" 
        };
        
        const fM = (n) => new Intl.NumberFormat('es-CL', { 
            style: 'currency', currency: 'CLP', minimumFractionDigits: 2 
        }).format(n);
        
        // Cabecera completa
        doc.setFontSize(16);
        doc.text("Reporte de Inventario Físico", 148.5, 15, { align: 'center' });
        doc.setFontSize(10);
        doc.text(`Bodega: ${meta.b} | Site: ${meta.s} | Fecha: ${meta.f} | Responsable: ${meta.r}`, 148.5, 22, { align: 'center' });
        
        const body = productosBase.map(p => {
            const f = parseFloat(localStorage.getItem(`inv-${p.codigo}-${p.lote}`)) || 0;
            const d = f - p.teorico;
            return [p.codigo, p.nombre, p.teorico.toFixed(2), f.toFixed(2), d.toFixed(2), fM(p.precio), fM(d * p.precio), fM(f * p.precio)];
        });

        doc.autoTable({
            startY: 28,
            head: [['Cód.', 'Producto', 'Sist.', 'Fís.', 'Dif.', 'Precio', 'Vr. Dif', 'Total Fís.']],
            body: body,
            styles: { fontSize: 7, halign: 'right' },
            columnStyles: { 0: {halign: 'left'}, 1: {halign: 'left'} },
            didParseCell: (data) => {
                // Colorear Diferencias (Col 4 y 6)
                if ([4, 6].includes(data.column.index)) {
                    const v = parseFloat(data.cell.raw.toString().replace(/[^0-9.-]/g, ''));
                    if (v < 0) data.cell.styles.textColor = [200, 0, 0];
                    else if (v > 0) data.cell.styles.textColor = [0, 128, 0];
                }
            }
        });

        const y = doc.lastAutoTable.finalY + 10;
        doc.setTextColor(0);
        // Totales sin decimales como pediste
        doc.text(`Total Sistema: ${new Intl.NumberFormat('es-CL', {style:'currency', currency:'CLP', maximumFractionDigits:0}).format(res.sistema)}`, 14, y);
        doc.text(`Total Físico:   ${new Intl.NumberFormat('es-CL', {style:'currency', currency:'CLP', maximumFractionDigits:0}).format(res.fisico)}`, 14, y + 7);
        
        // Diferencia y Porcentaje con color
        const color = res.difPesos < 0 ? [200, 0, 0] : [0, 128, 0];
        doc.setTextColor(color[0], color[1], color[2]);
        doc.setFont(undefined, 'bold');
        doc.text(`DIFERENCIA: ${fM(res.difPesos)}  (${res.porcentaje}%)`, 14, y + 14);
        
        doc.save(`Reporte_${meta.b}.pdf`);
    } catch (e) { alert("Error al generar PDF."); }
}

function formatearMoneda(v) {
    return new Intl.NumberFormat('es-CL', { style: 'currency', currency: 'CLP', maximumFractionDigits: 0 }).format(v);
}

function actualizarBarraProgreso() {
    const total = productosBase.length;
    if (total === 0) return;

    // Contamos cuántos productos tienen un valor físico ingresado en la memoria
    const contados = productosBase.filter(p => {
        const valor = localStorage.getItem(`inv-${p.codigo}-${p.lote}`);
        return valor !== null && valor !== ""; 
    }).length;

    const porc = Math.round((contados / total) * 100);
    const bar = document.getElementById('progress-bar');
    
    if (bar) {
        bar.style.width = porc + "%"; 
        bar.innerText = porc + "%";   
        
        // Color dinámico: Naranja si está en proceso, Verde si está al 100%
        if (porc < 100) {
            bar.style.backgroundColor = "#f39c12"; 
        } else {
            bar.style.backgroundColor = "#27ae60"; 
        }
    }
}

function filtrarProductos() {
    const t = document.getElementById('search').value.toLowerCase();
    const filtrados = productosBase.filter(p => 
        p.nombre.toLowerCase().includes(t) || 
        p.codigo.toLowerCase().includes(t) || 
        p.lote.toLowerCase().includes(t)
    );
    renderizarTarjetas(filtrados);
}

function limpiarTodo() {
    if(confirm("¿Seguro que quieres borrar todo el conteo actual?")) {
        // Borramos solo los datos del inventario, no la configuración (meta)
        productosBase.forEach(p => {
            localStorage.removeItem(`inv-${p.codigo}-${p.lote}`);
        });
        location.reload();
    }
}

function toggleDarkMode() {
    const body = document.body;
    const btn = document.getElementById('dark-mode-btn');
    
    body.classList.toggle('dark-mode');
    
    if (body.classList.contains('dark-mode')) {
        localStorage.setItem('theme', 'dark');
        btn.innerText = "☀️ Light";
    } else {
        localStorage.setItem('theme', 'light');
        btn.innerText = "🌙 Dark";
    }
}

// Al cargar la página, revisar si ya estaba en modo oscuro
window.addEventListener('DOMContentLoaded', () => {
    if (localStorage.getItem('theme') === 'dark') {
        document.body.classList.add('dark-mode');
        document.getElementById('dark-mode-btn').innerText = "☀️ Light";
    }
});