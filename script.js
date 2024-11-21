const imageUpload = document.getElementById("image-upload");
const previewContainer = document.getElementById("image-preview");

let selectedImages = [];
imageUpload.addEventListener("change", function (event) {
  previewContainer.innerHTML = "";
  selectedImages = [];

  const files = Array.from(event.target.files);

  files.forEach((file, index) => {
    const reader = new FileReader();

    reader.onload = function (e) {
      const imageContainer = document.createElement("div");
      imageContainer.classList.add("image-container");

      const img = document.createElement("img");
      img.src = e.target.result;
      img.alt = file.name;

      const removeButton = document.createElement("button");
      removeButton.innerHTML = '<i class="fas fa-times"></i>';
      removeButton.classList.add("remove-button");

      removeButton.onclick = () => {
        selectedImages = selectedImages.filter((_, i) => i !== index);
        imageContainer.remove();
      };

      imageContainer.appendChild(img);
      imageContainer.appendChild(removeButton);
      previewContainer.appendChild(imageContainer);

      selectedImages.push({
        name: file.name,
        file: file,
        url: e.target.result,
      });
    };

    reader.readAsDataURL(file);
  });
});

const optionSelect = document.getElementById("options");
const descriptionColumn = document.getElementById("description-column");

let optionDescriptions = {};
let optionsData = [];

optionSelect.addEventListener("change", () => {
  const selectedOptionText =
    optionSelect.options[optionSelect.selectedIndex]?.textContent;

  descriptionColumn.innerHTML = "";
  if (optionSelect.value) {
    let option = optionsData.find((opt) => opt.name === selectedOptionText);
    if (!option) {
      option = { name: selectedOptionText, description: "" };
      optionsData.push(option);
    }
    const label = document.createElement("label");
    label.textContent = `Descripción:`;

    const textarea = document.createElement("textarea");
    textarea.rows = 5;
    textarea.value = option.description;
    textarea.addEventListener("input", (e) => {
      option.description = e.target.value;
    });

    const deleteButton = document.createElement("button");
    deleteButton.type = "button";
    deleteButton.textContent = "Eliminar";
    deleteButton.className = "generate-button";
    deleteButton.addEventListener("click", () => {
      optionsData = optionsData.filter(
        (opt) => opt.name !== selectedOptionText
      );
      optionSelect.value = "";
      descriptionColumn.innerHTML =
        "<p class='placeholder'>Seleccione una opción para editar su descripción</p>";
    });

    descriptionColumn.appendChild(label);
    descriptionColumn.appendChild(textarea);
    descriptionColumn.appendChild(deleteButton);
  } else {
    descriptionColumn.innerHTML =
      "<p class='placeholder'>Seleccione una opción para editar su descripción</p>";
  }
});

const generateWordButton = document.getElementById("generate-word");
const generateWordButtonTable = document.getElementById("generate-table");
const {
  Document,
  AlignmentType,
  LevelFormat,
  UnderlineType,
  Packer,
  HeadingLevel,
  Tab,
  TabStopPosition,
  TabStopType,
  Paragraph,
  ImageRun,
  Table,
  TableCell,
  TableRow,
  TextRun,
  Alignment,
  BorderStyle,
  VerticalAlign,
  WidthType,
} = window.docx;

generateWordButton.addEventListener("click", function () {
  const children = [
    new Paragraph({
      alignment: "center",
      children: [new TextRun({ text: "REGISTRO FOTOGRAFICO", bold: true, size: 22 })],
      heading: HeadingLevel.TITLE,
      spacing: { before: 500, after: 600 },
    }),
  ];
  if (selectedImages.length === 0) {
    alert("Por favor, selecciona al menos una imagen.");
    return;
  }
  
  children.push(new Paragraph({ spacing: { after: 50 } }));
  for (let i = 0; i < selectedImages.length; i++) {
    children.push(
      gettable(
        selectedImages[i].url,
        selectedImages[i].name,
        "Descripcion de la imagen"
      )
    );
    if (i < selectedImages.length - 1) {
      children.push(new Paragraph({ spacing: { after: 200 } }));
    }
  }
  const doc = new Document({
    styles: {
      default: {
        heading1: {
          run: {
            size: "11pt",
            bold: true,
          },
          paragraph: {
            spacing: {
              after: 60,
            },
          },
        },
        heading2: {
          run: {
            bold: true,
          },
          paragraph: {
            spacing: {
              before: 120,
              after: 60,
            },
          },
        },
        document: {
          run: {
            size: "11pt",
            font: "Calibri",
          },
        },
      },
      paragraphStyles: [
        {
          id: "tabulated",
          name: "Tabulated Text",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          paragraph: {
            indent: { left: 360 },
            spacing: { after: 200 },
          },
        },
        {
          id: "table_header",
          name: "table_header",
          basedOn: "Normal",
          next: "Normal",
          run: {
            size: "9pt",
            bold: true,
          },
        },
        {
          id: "table_text",
          name: "table_text",
          basedOn: "Normal",
          next: "Normal",
          run: {
            size: "9pt",
          },
        },
        {
          id: "wellSpaced",
          name: "Well Spaced",
          basedOn: "Normal",
          quickFormat: true,
          paragraph: {
            spacing: {
              line: 276,
              before: 20 * 72 * 0.1,
              after: 20 * 72 * 0.05,
            },
          },
        },
        {
          id: "strikeUnderline",
          name: "Strike Underline",
          basedOn: "Normal",
          quickFormat: true,
          run: {
            strike: true,
            underline: {
              type: UnderlineType.SINGLE,
            },
          },
        },
      ],
    },
    numbering: {
      config: [
        {
          reference: "numbering",
          levels: [
            {
              level: 0,
              format: LevelFormat.DECIMAL,
              text: "%1.",
              alignment: "left",
              style: {
                paragraph: {
                  indent: { left: 200, hanging: 200 },
                  spacing: { after: 200 },
                },
              },
            },
            {
              level: 1,
              format: LevelFormat.DECIMAL,
              text: "%1.%2.",
              alignment: "left",
              style: {
                paragraph: {
                  indent: { left: 400, hanging: 100 },
                  spacing: { after: 50 },
                },
              },
            },
          ],
        },
      ],
    },
    sections: [{ children }],
  });

  Packer.toBlob(doc)
    .then((blob) => {
      saveAs(blob, "registro-fotografico.docx");
    })
    .catch((error) => {
      console.error("Error al generar el documento:", error);
    });
});

generateWordButtonTable.addEventListener("click", function () {
  const currentDate = new Date();
  const formattedDate = `${currentDate.toLocaleDateString("es-ES", {
    year: "numeric",
    month: "long",
    day: "numeric",
  })}`;
  const children = [
    new Paragraph({
      alignment: "center",
      children: [new TextRun({ text: "INFORME N° 1", bold: true, size: 22 })],
      heading: HeadingLevel.TITLE,
      spacing: { before: 500, after: 600 },
    }),
    new Paragraph({
      tabStops: [
        {
          type: TabStopType.RIGHT,
          position: TabStopPosition.MAX,
        },
      ],
      children: [
        new TextRun("Los Olivos, " + formattedDate),
        new TextRun({
          children: [new Tab(), "Expediente: 202200057119"],
          alignment: Alignment.right,
        }),
      ],
      alignment: Alignment.LEFT,
      thematicBreak: true,
      spacing: { after: 200 },
    }),
    new Paragraph({
      tabStops: [{ position: 1000, alignment: "left" }],
      indent: {
        left: 1500,
        hanging: 1500,
      },
      spacing: { after: 200 },
      children: [
        new TextRun({ text: "A\t:\tOficina Regional Lima Norte - Osinergmin" }),
      ],
    }),
    new Paragraph({
      tabStops: [{ position: 1000, alignment: "left" }],
      indent: {
        left: 1500,
        hanging: 1500,
      },
      spacing: { after: 200 },
      children: [
        new TextRun({ text: "De\t:\tIng. Paolo Minaya Ccaihuare" }),
        new TextRun({
          text: "CONSORCIO EVALUACIÓN Y SUPERVISIÓN EN ENERGIA EIRL, CALUSS S.A.C., CONSULTORÍA & SERVICIOS EN HIDROCARBUROS Y MINERÍA S.A.C.",
          break: 1,
        }),
      ],
    }),
    new Paragraph({
      tabStops: [{ position: 1000, alignment: "left" }],
      indent: {
        left: 1500,
        hanging: 1500,
      },
      children: [
        new TextRun({
          text: "Asunto\t:\tFiscalización de Condiciones de Seguridad en la Estación de Servicios con ficha de registro N° 8082-107-270918 que tiene como titular a ESTACIÓN LOS JARDINES E.I.R.L.",
        }),
      ],
    }),
    new Paragraph({
      spacing: { before: 200, after: 300 },
      thematicBreak: true,
    }),
    new Paragraph({
      text: "OBJETIVO",
      heading: HeadingLevel.HEADING_1,
      numbering: {
        reference: "numbering",
        level: 0,
      },
    }),
    new Paragraph({
      text: "Realizar la fiscalización de Condiciones de Seguridad del establecimiento ESTACIÓN LOS JARDINES E.I.R.L. con ficha de registro vigente.",
      style: "tabulated",
    }),
    new Paragraph({
      text: "ANTECEDENTES",
      heading: HeadingLevel.HEADING_1,
      numbering: {
        reference: "numbering",
        level: 0,
      },
    }),
    new Paragraph({
      text: "Con fecha 27 de marzo de 2022, la Oficina Regional Lima Norte asignó el expediente Nº 202000181786 y la carta línea N° 743647-1 a la empresa supervisora CONSORCIO EVALUACIÓN Y SUPERVISIÓN EN ENERGIA EIRL, CALUSS S.A.C., CONSULTORÍA & SERVICIOS EN HIDROCARBUROS Y MINERÍA S.A.C., el mismo que derivó a su supervisor el ING. PAOLO ISIDORO MINAYA CCAIHUARE, para efectuar la visita de al establecimiento asignado.",
      numbering: {
        reference: "numbering",
        level: 1,
      },
    }),
    new Paragraph({
      text: "Con fecha 31 de marzo de 2022, se realizó la visita de fiscalización a la Estación de Servicios mencionada líneas arriba.",
      numbering: {
        reference: "numbering",
        level: 1,
      },
    }),
    new Paragraph({
      text: "BASE LEGAL",
      heading: HeadingLevel.HEADING_1,
      numbering: {
        reference: "numbering",
        level: 0,
      },
    }),
    new Paragraph({
      text: "Reglamento del Registro de Hidrocarburos, aprobado por Resolución de Consejo Directivo N° 191-2011 OS/CD y sus modificatorias.",
      numbering: {
        reference: "numbering",
        level: 1,
      },
    }),
    new Paragraph({
      text: "Decreto Supremo N° 01-94-EM.",
      numbering: {
        reference: "numbering",
        level: 1,
      },
    }),
    new Paragraph({
      text: "Decreto Supremo N° 027-94-EM.",
      numbering: {
        reference: "numbering",
        level: 1,
      },
    }),
    new Paragraph({
      text: "ANÁLISIS Y RESULTADOS",
      heading: HeadingLevel.HEADING_1,
      numbering: {
        reference: "numbering",
        level: 0,
      },
    }),
    new Paragraph({
      text: "Se realizó la visita de fiscalización para verificar la Declaración Jurada presentada por el Agente Fiscalizado en el sistema PVO de Osinergmin. La Declaración Jurada fue presentada el 30 de octubre de 2021. Con base en ello, se realizó la evaluación si el Agente Fiscalizado cumplía con lo mencionado en dicho documento.",
      style: "tabulated",
    }),
    new Paragraph({
      text: "Evaluación del PDJ",
      heading: HeadingLevel.HEADING_2,
      numbering: {
        reference: "numbering",
        level: 1,
      },
    }),
    new Paragraph({
      text: "Durante la visita de fiscalización del 31 de marzo de 2022, se encontraron los siguientes incumplimientos:",
      style: "tabulated",
    }),
  ];
  children.push(new Paragraph({ spacing: { after: 50 } }));
  if (optionsData.length > 0) {
    children.push(createTableWithHeaders(optionsData));
  }
  children.push(new Paragraph({ spacing: { after: 50 } }));
  children.push(
    new Paragraph({
      text: "Otros hallazgos",
      heading: HeadingLevel.HEADING_2,
      numbering: {
        reference: "numbering",
        level: 1,
      },
    })
  );
  children.push(
    new Paragraph({
      text: "Durante la visita de fiscalización se encontró que en la isla N°4 el segundo equipo de GLP no se tenía en uso, debido a que no cumple con las distancias de 6.1m respecto a los puntos de varillaje de los diferentes tanques de Combustible líquido. El encargado nos informó que están cambiando las conexiones de varillaje para comenzar a utilizarlo lo antes posible.",
      style: "tabulated",
    })
  );
  children.push(
    new Paragraph({
      text: "Además, se encontró que la ficha de registro vigente, no coincide con lo encontrado en la visita de fiscalización respecto a los tanques y las disposiciones correspondientes. Además, ahora ya retiraron el tanque que lo tenían sin producto. Se ha recomendado que lo actualicen los datos correspondientes.",
      style: "tabulated",
    })
  );
  children.push(
    new Paragraph({
      text: "CONCLUSIÓN",
      heading: HeadingLevel.HEADING_1,
      numbering: {
        reference: "numbering",
        level: 0,
      },
    })
  );
  children.push(
    new Paragraph({
      text: "El Agente Fiscalizado tiene observaciones que van contra la normativa vigente. Se deben mantener dichas observaciones hasta que el Agente Fiscalizado los subsane y mientras enviar al área legal para iniciar el PAS.",
      style: "tabulated",
    })
  );
  children.push(
    new Paragraph({
      text: "RECOMENDACIÓN",
      heading: HeadingLevel.HEADING_1,
      numbering: {
        reference: "numbering",
        level: 0,
      },
    })
  );
  children.push(
    new Paragraph({
      text: "Se recomienda enviar para comenzar el Inicio de PAS debido a que durante la visita de fiscalización se encontraron incumplimientos en el establecimiento asignado.",
      style: "tabulated",
    })
  );
  const doc = new Document({
    styles: {
      default: {
        heading1: {
          run: {
            size: "11pt",
            bold: true,
          },
          paragraph: {
            spacing: {
              after: 60,
            },
          },
        },
        heading2: {
          run: {
            bold: true,
          },
          paragraph: {
            spacing: {
              before: 120,
              after: 60,
            },
          },
        },
        document: {
          run: {
            size: "11pt",
            font: "Calibri",
          },
        },
      },
      paragraphStyles: [
        {
          id: "tabulated",
          name: "Tabulated Text",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          paragraph: {
            indent: { left: 360 },
            spacing: { after: 200 },
          },
        },
        {
          id: "table_header",
          name: "table_header",
          basedOn: "Normal",
          next: "Normal",
          run: {
            size: "9pt",
            bold: true,
          },
        },
        {
          id: "table_text",
          name: "table_text",
          basedOn: "Normal",
          next: "Normal",
          run: {
            size: "9pt",
          },
        },
        {
          id: "wellSpaced",
          name: "Well Spaced",
          basedOn: "Normal",
          quickFormat: true,
          paragraph: {
            spacing: {
              line: 276,
              before: 20 * 72 * 0.1,
              after: 20 * 72 * 0.05,
            },
          },
        },
        {
          id: "strikeUnderline",
          name: "Strike Underline",
          basedOn: "Normal",
          quickFormat: true,
          run: {
            strike: true,
            underline: {
              type: UnderlineType.SINGLE,
            },
          },
        },
      ],
    },
    numbering: {
      config: [
        {
          reference: "numbering",
          levels: [
            {
              level: 0,
              format: LevelFormat.DECIMAL,
              text: "%1.",
              alignment: "left",
              style: {
                paragraph: {
                  indent: { left: 200, hanging: 200 },
                  spacing: { after: 200 },
                },
              },
            },
            {
              level: 1,
              format: LevelFormat.DECIMAL,
              text: "%1.%2.",
              alignment: "left",
              style: {
                paragraph: {
                  indent: { left: 400, hanging: 100 },
                  spacing: { after: 50 },
                },
              },
            },
          ],
        },
      ],
    },
    sections: [{ children }],
  });

  Packer.toBlob(doc)
    .then((blob) => {
      saveAs(blob, "informe-observaciones.docx");
    })
    .catch((error) => {
      console.error("Error al generar el documento:", error);
    });
});

function gettable(img, nombreImagen, descripcion) {
  return new Table({
    alignment: "center",
    rows: [
      new TableRow({
        children: [
          new TableCell({
            margins: { top: 100, bottom: 100, left: 200, right: 200 },
            children: [
              new Paragraph({
                children: [
                  new ImageRun({
                    data: img,
                    transformation: {
                      width: 500,
                      height: 400,
                    },
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
      new TableRow({
        margins: 5,
        children: [
          new TableCell({
            margins: { top: 100, bottom: 100, left: 100, right: 100 },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: nombreImagen.split(".").slice(0, -1).join("."),
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
  const lines = descripcion
    .trim()
    .split("\n")
    .filter((line) => line.trim() !== "");
  const textRuns = lines.map(
    (line, i) =>
      new TextRun({
        text: line,
        break: i == 0 ? 0 : 1,
      })
  );
  return new TableRow({
    children: [
      new TableCell({
        margins: { left: 100, right: 100 },
        verticalAlign: "center",
        children: [
          new Paragraph({
            text: (index + 1).toString(),
            alignment: Alignment.CENTER,
            style: "table_text",
          }),
        ],
      }),
      new TableCell({
        margins: { left: 100, right: 100 },
        children: [new Paragraph({ children: textRuns, style: "table_text" })],
      }),
      new TableCell({
        margins: { left: 100, right: 100 },
        children: [new Paragraph({ text: title, style: "table_text" })],
      }),
    ],
  });
}

function createTableWithHeaders(Opciones) {
  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      new TableCell({
        margins: { top: 200, bottom: 200, left: 100, right: 100 },
        children: [
          new Paragraph({
            text: "N°",
            alignment: Alignment.CENTER,
            style: "table_header",
          }),
        ],
        verticalAlign: VerticalAlign.CENTER,
        shading: { fill: "D3D3D3" },
      }),
      new TableCell({
        margins: { top: 200, bottom: 200, left: 100, right: 100 },
        children: [
          new Paragraph({
            text: "Incumplimiento verificado",
            alignment: Alignment.CENTER,
            style: "table_header",
          }),
        ],
        verticalAlign: VerticalAlign.CENTER,
        shading: { fill: "D3D3D3" },
      }),
      new TableCell({
        margins: { top: 200, bottom: 200, left: 100, right: 100 },
        children: [
          new Paragraph({
            text: "Base Legal",
            alignment: Alignment.CENTER,
            style: "table_header",
          }),
        ],
        verticalAlign: VerticalAlign.CENTER,
        shading: { fill: "D3D3D3" },
      }),
    ],
  });

  const rows = [headerRow];
  for (let i = 0; i < Opciones.length; i++) {
    rows.push(
      createEvaluacionTable(i, Opciones[i].name, Opciones[i].description)
    );
  }

  return new Table({
    alignment: "center",
    rows: rows,
    width: {
      size: 100,
      type: WidthType.PERCENTAGE,
    },
  });
}
