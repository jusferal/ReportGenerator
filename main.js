const fs = require("fs");
const path = require("path");
const probe = require('probe-image-size');
const { Document, Packer, Paragraph, ImageRun, TextRun,Table, TableCell, TableRow, Tab, Alignment, BorderStyle } = require("docx");
const { Console } = require("console");

const imagesFolderPath = "./imagenes";

// Crear un documento con una sola sección
function gettable(ruta, descripcion){
    const nombreImagen = path.parse(ruta).name;
    const imageBuffer = fs.readFileSync(ruta);
    const result = probe.sync(imageBuffer);
    return new Table({
        alignment:"center",
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        margins: { top: 100, bottom: 100, left: 200, right: 200 },
                        children: [new Paragraph({
                            children: [
                                new ImageRun({
                                    data: fs.readFileSync(ruta),
                                    transformation: {
                                        width: 500,
                                        height: 400,
                                    }
                                }),
                            ],
                        }),],
                    }),
                ],
            }),
            new TableRow({
                margins:5,
                children: [
                    new TableCell({
                        margins: { top: 100, bottom: 100, left: 100, right: 100 },
                        children: [
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: nombreImagen,
                                        bold: true,
                                    }),
                                ],
                            }),
                            new Paragraph(descripcion),
                        ],
                    }),
                ],
            }),
        ],
    });
}

function createEvaluacionTable(index, title, descripcion) {
    return new TableRow({
        children: [
            new TableCell({
                children: [new Paragraph((index + 1).toString())], // Numeración
            }),
            new TableCell({
                children: [new Paragraph(descripcion)], // Título
            }),
            new TableCell({
                children: [new Paragraph(title)], // Título
            }),
        ],
    });
}

function createTableWithHeaders(titles, descriptions) {
    const headerRow = new TableRow({
        children: [
            new TableCell({ children: [new Paragraph("N°")], alignment: Alignment.CENTER,shading: { fill: "D3D3D3" },  }),
            new TableCell({ children: [new Paragraph("Inclumplimiento verificado")], alignment: Alignment.CENTER,shading: { fill: "D3D3D3" }, }),
            new TableCell({ children: [new Paragraph("Base Legal")], alignment: Alignment.CENTER,shading: { fill: "D3D3D3" }, }),
        ],
    });

    const rows = [headerRow];
    for (let i = 0; i < titles.length; i++) {
        rows.push(createEvaluacionTable(i, descriptions[i], titles[i]));
    }

    return new Table({
        alignment: Alignment.CENTER,
        rows: rows,
    });
}


const children = [
    new Paragraph({
        children: [
            new TextRun({ text: "INFORME N° «osinumero»", bold: true, size: 22, alignment:Alignment.CENTER }),
        ],
        alignment: Alignment.CENTER,
    }),
    new Paragraph({
        children: [
            new TextRun("Los Olivos, 13 de abril de 2022"),
            new TextRun("\t\t\t\tExpediente: 202200057119"),
        ],
        alignment: Alignment.LEFT,
    }),
    new Paragraph({
        border: {
            bottom: {
                color: "000000", // Color de la línea (negro)
                space: 3,       // Espacio entre el borde y el contenido
                value: BorderStyle.SINGLE,
                size: 6,        // Grosor del borde
            },
        },
        children: [],
    }),
    new Paragraph({
        children: [
            new TextRun({ text: "A:\tOficina Regional Lima Norte - Osinergmin", bold: true }),
        ],
    }),
    new Paragraph({
        children: [
            new TextRun({ text: "De:\tIng. Paolo Minaya Ccaihuare", bold: true }),
            new TextRun("\nCONSORCIO EVALUACIÓN Y SUPERVISIÓN EN ENERGIA EIRL, CALUSS S.A.C., CONSULTORÍA & SERVICIOS EN HIDROCARBUROS Y MINERÍA S.A.C."),
        ],
    }),
    new Paragraph({
        children: [
            new TextRun({
                text: "Asunto: Fiscalización de Condiciones de Seguridad en la Estación de Servicios con ficha de registro N° 8082-107-270918 que tiene como titular a ESTACIÓN LOS JARDINES E.I.R.L.",
                bold: true,
            }),
        ],
    }),
    // Encabezados para cada sección
    new Paragraph({
        text: "OBJETIVO",
        heading: "Heading2",
    }),
    new Paragraph(
        "Realizar la fiscalización de Condiciones de Seguridad del establecimiento ESTACIÓN LOS JARDINES E.I.R.L. con ficha de registro vigente."
    ),
    new Paragraph({
        text: "ANTECEDENTES",
        heading: "Heading2",
    }),
    new Paragraph(
        "Con fecha 27 de marzo de 2022, la Oficina Regional Lima Norte asignó el expediente Nº 202000181786..."
    ),
    new Paragraph({
        text: "BASE LEGAL",
        heading: "Heading2",
    }),
    new Paragraph(
        "Reglamento del Registro de Hidrocarburos, aprobado por Resolución de Consejo Directivo N° 191-2011 OS/CD..."
    ),
    new Paragraph({
        text: "ANÁLISIS Y RESULTADOS",
        heading: "Heading2",
    }),
    new Paragraph(
        "Se realizó la visita de fiscalización para verificar la Declaración Jurada presentada por el Agente Fiscalizado..."
    ),
    new Paragraph({
        alignment:"center",
        children: [
            new TextRun({text:"REGISTRO FOTOGRAFICO", bold: true}),
        ],
    }),
];

const images = fs.readdirSync(imagesFolderPath).filter(file =>
    file.endsWith('.jpeg') || file.endsWith('.png')|| file.endsWith('.jpg')
);
images.forEach((imageName, index) => {
    const imagePath = path.join(imagesFolderPath, imageName);
    const descripcion = `Vista de la imagen ${imageName}`; 
    children.push(gettable(imagePath, descripcion));
    if (index < images.length - 1) { 
        children.push(new Paragraph({ spacing: { after: 200 } }));
    }
});

const titles = ["Título 1", "Título 2", "Título 3"];
const descriptions = ["Descripción 1", "Descripción 2", "Descripción 3"];
children.push(new Paragraph({ spacing: { after: 200 } }));
children.push(createTableWithHeaders(titles, descriptions));

const doc = new Document({
    sections: [
        {
            properties: {},
            children: children,
        },
    ],
});


Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("Document.docx", buffer);
    console.log("Documento creado exitosamente: My Document.docx");
});
