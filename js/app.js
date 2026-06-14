const usuarios = [
    "Cindy", "Packy", "Wendy", "Lennon", "Yulian", "Jimmy", "Erick", "Erik", "Bere"
];

async function cargarExcel() {
    try {
        const respuesta = await fetch("./data/Quiniela.xlsx");
        if (!respuesta.ok) throw new Error("No se pudo cargar el archivo");
        const archivo = await respuesta.arrayBuffer();
        const workbook = XLSX.read(archivo);

        cargarRanking(workbook);
        cargarUsuarios(workbook);
    } catch (error) {
        console.error("Error al cargar la quiniela:", error);
        document.getElementById("top3").innerHTML = "<p>Error al cargar datos. Verifica el archivo.</p>";
    }
}

function cargarRanking(workbook) {
    const hojaRanking = workbook.Sheets["Ranking"];
    const datos = XLSX.utils.sheet_to_json(hojaRanking, { header: 1 });

    const tbody = document.getElementById("rankingBody");
    const top3 = document.getElementById("top3");

    tbody.innerHTML = "";
    top3.innerHTML = "";

    let ranking = [];

    for (let fila = 4; fila <= 12; fila++) {
        const posicion = datos[fila]?.[3];
        const usuario = datos[fila]?.[4];
        const puntos = datos[fila]?.[5];

        if (!usuario) continue;

        ranking.push({ posicion, usuario, puntos });

        let medalla = "";
        if (posicion == 1) medalla = "🥇";
        else if (posicion == 2) medalla = "🥈";
        else if (posicion == 3) medalla = "🥉";

        tbody.innerHTML += `
            <tr>
                <td>${medalla} ${posicion}</td>
                <td>${usuario}</td>
                <td>${puntos}</td>
            </tr>
        `;
    }

    if (ranking.length >= 3) {
        top3.innerHTML = `
            <div class="tarjeta-top oro">🥇<br>${ranking[0].usuario}<br>${ranking[0].puntos} pts</div>
            <div class="tarjeta-top plata">🥈<br>${ranking[1].usuario}<br>${ranking[1].puntos} pts</div>
            <div class="tarjeta-top bronce">🥉<br>${ranking[2].usuario}<br>${ranking[2].puntos} pts</div>
        `;
    }
}

function cargarUsuarios(workbook) {
    const contenedor = document.getElementById("usuarios");
    contenedor.innerHTML = "";

    usuarios.forEach(usuario => {
        const div = document.createElement("div");
        div.className = "usuario";
        div.textContent = usuario;
        div.onclick = () => mostrarPronosticos(workbook, usuario);
        contenedor.appendChild(div);
    });
}

function mostrarPronosticos(workbook, usuario) {
    const hojaUsuario = workbook.Sheets[usuario];
    const hojaResultados = workbook.Sheets["Resultados"]; // 1. Cargamos la hoja de Resultados Reales

    if (!hojaUsuario || !hojaResultados) {
        document.getElementById("detalleUsuario").innerHTML = `<p>Error: No se encontraron los datos necesarios.</p>`;
        return;
    }

    const datosUsuario = XLSX.utils.sheet_to_json(hojaUsuario, { header: 1 });
    const datosResultados = XLSX.utils.sheet_to_json(hojaResultados, { header: 1 });

    let totalPuntos = 0;
    let html = `<h3>👤 Pronósticos de ${usuario}</h3>`;
    
    // Modificamos la cabecera de la tabla para incluir el "Resultado Real"
    let tabla = `
        <table>
            <thead>
                <tr>
                    <th>Local</th>
                    <th>Pronóstico</th>
                    <th>Resultado Real</th>
                    <th>Visitante</th>
                    <th>Puntos</th>
                </tr>
            </thead>
            <tbody>
    `;

    // Asumimos que la hoja de Resultados tiene la misma estructura de filas que la del usuario
    for (let i = 3; i < datosUsuario.length; i++) {
        const local = datosUsuario[i]?.[4];
        const golLocal = datosUsuario[i]?.[5];
        const golVisitante = datosUsuario[i]?.[6];
        const visitante = datosUsuario[i]?.[7];
        const puntos = Number(datosUsuario[i]?.[9] || 0);

        if (!local || !visitante || ["Local", "Casa", "Fecha"].includes(local)) {
            continue;
        }

        // 2. Extraemos el resultado real correspondiente a la misma fila
        const realGolLocal = datosResultados[i]?.[5];
        const realGolVisitante = datosResultados[i]?.[6];

        // Formateamos el resultado real (si aún no se juega, ponemos un guión o vacío)
        let resultadoRealTexto = "-";
        if (realGolLocal !== undefined && realGolVisitante !== undefined) {
            resultadoRealTexto = `${realGolLocal} - ${realGolVisitante}`;
        }

        totalPuntos += puntos;

        let clase = "puntos-0";
        let emoji = "❌";

        if (puntos === 2) {
            clase = "puntos-2";
            emoji = "✅";
        } else if (puntos === 1) {
            clase = "puntos-1";
            emoji = "🟨";
        }

        // Agregamos la columna del resultado real a la fila de la tabla
        tabla += `
            <tr>
                <td>${local}</td>
                <td>${golLocal} - ${golVisitante}</td>
                <td style="font-weight: bold; color: #07829b;">${resultadoRealTexto}</td>
                <td>${visitante}</td>
                <td class="${clase}">${emoji} ${puntos}</td>
            </tr>
        `;
    }

    tabla += `</tbody></table>`;

    html += `
        <p style="margin:10px 0 20px 0;">
            🏆 <strong>Total de puntos calculados: ${totalPuntos}</strong>
        </p>
    `;
    html += tabla;

    document.getElementById("detalleUsuario").innerHTML = html;
}

// Inicializar la carga
cargarExcel();