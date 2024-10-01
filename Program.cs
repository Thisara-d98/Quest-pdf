using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using QuestPDF.Previewer;

using PdfDocument = QuestPDF.Fluent.Document;

QuestPDF.Settings.License = LicenseType.Community;

var fileName = "working-sheet.pdf";

PdfDocument.Create(container =>
{
    container.Page(page =>
    {
        page.Size(PageSizes.A4);
        page.Margin(1, Unit.Centimetre);
        page.PageColor(Colors.White);
        page.DefaultTextStyle(x => x.FontSize(10));

        // Header for Consignee and Exporter
        page.Header()
            .Column(column =>
            {
                column.Spacing(10);

                column.Item().Row(row =>
                {
                    row.RelativeItem().Text("Consignee:").SemiBold().FontSize(12);
                    row.RelativeItem().Text(""); // Empty space
                    row.RelativeItem().Text("Exporter:").SemiBold().FontSize(12).AlignCenter();
                });

                // Space below Consignee and Exporter
                column.Item().Text("");

                // "Working Sheet" at the center
                column.Item().AlignCenter().Text("Working Sheet")
                    .Italic().Underline().FontSize(14);
            });

        // Job No and Currency in the next row
        page.Content()
            .PaddingVertical(1, Unit.Centimetre)
            .Column(x =>
            {
                x.Spacing(10);

                // Job No and Currency
                x.Item().Row(row =>
                {
                    row.RelativeItem().Text("Job No: 20249351").FontSize(10);
                    row.RelativeItem().Text(""); // Empty space
                    row.RelativeItem().Text("Currency: JPY").AlignRight().FontSize(10);
                });

                // Date and Currency Rate in the next row
                x.Item().Row(row =>
                {
                    row.RelativeItem().Text("Date: Thursday, 22 Aug 2024 - 11:04:48 AM").FontSize(10);
                    row.RelativeItem().Text(""); // Empty space
                    row.RelativeItem().Text("Currency Rate: 1.1833").AlignRight().FontSize(10);
                });

                // Table starts here
                x.Item().Table(table =>
                {
                    table.ColumnsDefinition(columns =>
                    {
                        for (int i = 0; i < 11; i++) // 11 columns
                        {
                            columns.RelativeColumn();
                        }
                    });

                    // ITEM 001
                    table.Cell().Element(CellStyleWithoutBorders).Text("ITEM: 001").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("40169990").Bold();
                    for (int i = 0; i < 9; i++) table.Cell().Element(CellStyleWithoutBorders).Text("");

                    // Header row for ITEM 001
                    table.Cell().Element(CellStyleWithoutBorders).Text("No").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Description").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Qty 1").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("MU Qty2").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("FOB").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Freight").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Ins").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Others").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("CIF").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Nett").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Pkgs").Bold();

                    // Data row for ITEM 001
                    table.Cell().Element(CellStyleWithoutBorders).Text("1");
                    table.Cell().Element(CellStyleWithoutBorders).Text("BELT V");
                    table.Cell().Element(CellStyleWithoutBorders).Text("1.47Kg");
                    table.Cell().Element(CellStyleWithoutBorders).Text("05 PCS");
                    table.Cell().Element(CellStyleWithoutBorders).Text("5645.00");
                    table.Cell().Element(CellStyleWithoutBorders).Text("621.64");
                    table.Cell().Element(CellStyleWithoutBorders).Text("8.96");
                    table.Cell().Element(CellStyleWithoutBorders).Text("597.64");
                    table.Cell().Element(CellStyleWithoutBorders).Text("6873.24");
                    table.Cell().Element(CellStyleWithoutBorders).Text("1.47");
                    table.Cell().Element(CellStyleWithoutBorders).Text("0.04");

                    table.Cell().Element(CellStyleWithoutBorders).Text("2");
                    table.Cell().Element(CellStyleWithoutBorders).Text("RUBBER NO. 3, ENG.M");
                    table.Cell().Element(CellStyleWithoutBorders).Text("18.70Kg");
                    table.Cell().Element(CellStyleWithoutBorders).Text("10 PCS");
                    table.Cell().Element(CellStyleWithoutBorders).Text("71760.00");
                    table.Cell().Element(CellStyleWithoutBorders).Text("7902.36");
                    table.Cell().Element(CellStyleWithoutBorders).Text("113.86");
                    table.Cell().Element(CellStyleWithoutBorders).Text("7597.33");
                    table.Cell().Element(CellStyleWithoutBorders).Text("87373.55");
                    table.Cell().Element(CellStyleWithoutBorders).Text("18.70");
                    table.Cell().Element(CellStyleWithoutBorders).Text("0.46");

                    table.Cell().Element(CellStyleWithoutBorders).Text("3");
                    table.Cell().Element(CellStyleWithoutBorders).Text("BELT V");
                    table.Cell().Element(CellStyleWithoutBorders).Text("7.99Kg");
                    table.Cell().Element(CellStyleWithoutBorders).Text("25 PCS");
                    table.Cell().Element(CellStyleWithoutBorders).Text("30685.00");
                    table.Cell().Element(CellStyleWithoutBorders).Text("3379.10");
                    table.Cell().Element(CellStyleWithoutBorders).Text("48.69");
                    table.Cell().Element(CellStyleWithoutBorders).Text("3248.66");
                    table.Cell().Element(CellStyleWithoutBorders).Text("37361.45");
                    table.Cell().Element(CellStyleWithoutBorders).Text("7.99");
                    table.Cell().Element(CellStyleWithoutBorders).Text("0.20");

                    table.Cell().Element(CellStyleWithoutBorders).Text("4");
                    table.Cell().Element(CellStyleWithoutBorders).Text("RUBBER NO.1 ENG");
                    table.Cell().Element(CellStyleWithoutBorders).Text("3.74Kg");
                    table.Cell().Element(CellStyleWithoutBorders).Text("05 PCS");
                    table.Cell().Element(CellStyleWithoutBorders).Text("14355.00");
                    table.Cell().Element(CellStyleWithoutBorders).Text("1580.80");
                    table.Cell().Element(CellStyleWithoutBorders).Text("22.78");
                    table.Cell().Element(CellStyleWithoutBorders).Text("1519.78");
                    table.Cell().Element(CellStyleWithoutBorders).Text("17478.36");
                    table.Cell().Element(CellStyleWithoutBorders).Text("3.74");
                    table.Cell().Element(CellStyleWithoutBorders).Text("0.09");

                    table.Cell().Element(CellStyleWithoutBorders).Text("15");
                    table.Cell().Element(CellStyleWithoutBorders).Text("RUBBER NO. 4 ENG.M");
                    table.Cell().Element(CellStyleWithoutBorders).Text("5.85Kg");
                    table.Cell().Element(CellStyleWithoutBorders).Text("04 PCS");
                    table.Cell().Element(CellStyleWithoutBorders).Text("22460.00");
                    table.Cell().Element(CellStyleWithoutBorders).Text("2473.35");
                    table.Cell().Element(CellStyleWithoutBorders).Text("35.64");
                    table.Cell().Element(CellStyleWithoutBorders).Text("2377.87");
                    table.Cell().Element(CellStyleWithoutBorders).Text("27346.86");
                    table.Cell().Element(CellStyleWithoutBorders).Text("5.85");
                    table.Cell().Element(CellStyleWithoutBorders).Text("0.15");

                    // Totals for ITEM 001 with top and bottom borders
                    table.Cell().Element(CellStyleWithBorders).Text("");
                    table.Cell().Element(CellStyleWithBorders).Text("");
                    table.Cell().Element(CellStyleWithBorders).Text("37.75Kg");
                    table.Cell().Element(CellStyleWithBorders).Text("49.00");
                    table.Cell().Element(CellStyleWithBorders).Text("144905.00");
                    table.Cell().Element(CellStyleWithBorders).Text("15957.25");
                    table.Cell().Element(CellStyleWithBorders).Text("229.93");
                    table.Cell().Element(CellStyleWithBorders).Text("15341.28");
                    table.Cell().Element(CellStyleWithBorders).Text("176433.46");
                    table.Cell().Element(CellStyleWithBorders).Text("37.75");
                    table.Cell().Element(CellStyleWithBorders).Text("0.94");

                    // ITEM 002
                    table.Cell().Element(CellStyleWithoutBorders).Text("ITEM: 002").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("84219920").Bold();
                    for (int i = 0; i < 9; i++) table.Cell().Element(CellStyleWithoutBorders).Text("");

                    // Header row for ITEM 002
                    table.Cell().Element(CellStyleWithoutBorders).Text("No").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Description").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Qty 1").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("MU Qty2").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("FOB").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Freight").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Ins").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Others").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("CIF").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Nett").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Pkgs").Bold();

                    // Data row for ITEM 002
                    table.Cell().Element(CellStyleWithoutBorders).Text("5");
                    table.Cell().Element(CellStyleWithoutBorders).Text("ELEMENT AIR CLEAN");
                    table.Cell().Element(CellStyleWithoutBorders).Text("4.18Kg");
                    table.Cell().Element(CellStyleWithoutBorders).Text("15 PCS");
                    table.Cell().Element(CellStyleWithoutBorders).Text("16050.00");
                    table.Cell().Element(CellStyleWithoutBorders).Text("1767.46");
                    table.Cell().Element(CellStyleWithoutBorders).Text("25.47");
                    table.Cell().Element(CellStyleWithoutBorders).Text("1699.23");
                    table.Cell().Element(CellStyleWithoutBorders).Text("19542.16");
                    table.Cell().Element(CellStyleWithoutBorders).Text("4.18");
                    table.Cell().Element(CellStyleWithoutBorders).Text("0.10");

                    table.Cell().Element(CellStyleWithBorders).Text("");
                    table.Cell().Element(CellStyleWithBorders).Text("");
                    table.Cell().Element(CellStyleWithBorders).Text("4.18Kg");
                    table.Cell().Element(CellStyleWithBorders).Text("15.00");
                    table.Cell().Element(CellStyleWithBorders).Text("16050.00");
                    table.Cell().Element(CellStyleWithBorders).Text("1767.46");
                    table.Cell().Element(CellStyleWithBorders).Text("25.47");
                    table.Cell().Element(CellStyleWithBorders).Text("1699.23");
                    table.Cell().Element(CellStyleWithBorders).Text("19542.16");
                    table.Cell().Element(CellStyleWithBorders).Text("4.18");
                    table.Cell().Element(CellStyleWithBorders).Text("0.10");

                    // ITEM 003
                    table.Cell().Element(CellStyleWithoutBorders).Text("ITEM: 003").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("85111000").Bold();
                    for (int i = 0; i < 9; i++) table.Cell().Element(CellStyleWithoutBorders).Text("");

                    // Header row for ITEM 003
                    table.Cell().Element(CellStyleWithoutBorders).Text("No").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Description").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Qty 1").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("MU Qty2").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("FOB").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Freight").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Ins").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Others").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("CIF").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Nett").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Pkgs").Bold();

                    // Data row for ITEM 003
                    table.Cell().Element(CellStyleWithoutBorders).Text("6");
                    table.Cell().Element(CellStyleWithoutBorders).Text("PLUG SPARK");
                    table.Cell().Element(CellStyleWithoutBorders).Text("96.00UN");
                    table.Cell().Element(CellStyleWithoutBorders).Text("21.79 PCS");
                    table.Cell().Element(CellStyleWithoutBorders).Text("83616.00");
                    table.Cell().Element(CellStyleWithoutBorders).Text("9207.97");
                    table.Cell().Element(CellStyleWithoutBorders).Text("132.67");
                    table.Cell().Element(CellStyleWithoutBorders).Text("8852.55");
                    table.Cell().Element(CellStyleWithoutBorders).Text("101809.19");
                    table.Cell().Element(CellStyleWithoutBorders).Text("21.78");
                    table.Cell().Element(CellStyleWithoutBorders).Text("0.53");

                    // Totals for ITEM 003 with top and bottom borders
                    table.Cell().Element(CellStyleWithBorders).Text(" ");
                    table.Cell().Element(CellStyleWithBorders).Text(" ");
                    table.Cell().Element(CellStyleWithBorders).Text("96.00UN");
                    table.Cell().Element(CellStyleWithBorders).Text("21.79 PCS");
                    table.Cell().Element(CellStyleWithBorders).Text("83616.00");
                    table.Cell().Element(CellStyleWithBorders).Text("9207.97");
                    table.Cell().Element(CellStyleWithBorders).Text("132.67");
                    table.Cell().Element(CellStyleWithBorders).Text("8852.55");
                    table.Cell().Element(CellStyleWithBorders).Text("101809.19");
                    table.Cell().Element(CellStyleWithBorders).Text("21.78");
                    table.Cell().Element(CellStyleWithBorders).Text("0.53");

                    // ITEM 004
                    table.Cell().Element(CellStyleWithoutBorders).Text("ITEM: 004").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("40161000").Bold();
                    for (int i = 0; i < 9; i++) table.Cell().Element(CellStyleWithoutBorders).Text("");

                    // Header row for ITEM 004
                    table.Cell().Element(CellStyleWithoutBorders).Text("No").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Description").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Qty 1").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("MU Qty2").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("FOB").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Freight").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Ins").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Others").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("CIF").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Nett").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Pkgs").Bold();

                    // Data row for ITEM 004
                    table.Cell().Element(CellStyleWithoutBorders).Text("7");
                    table.Cell().Element(CellStyleWithoutBorders).Text("TENSIONER BELT V");
                    table.Cell().Element(CellStyleWithoutBorders).Text("7.59Kg");
                    table.Cell().Element(CellStyleWithoutBorders).Text("04 PCS");
                    table.Cell().Element(CellStyleWithoutBorders).Text("29124.00");
                    table.Cell().Element(CellStyleWithoutBorders).Text("3207.20");
                    table.Cell().Element(CellStyleWithoutBorders).Text("46.21");
                    table.Cell().Element(CellStyleWithoutBorders).Text("3083.40");
                    table.Cell().Element(CellStyleWithoutBorders).Text("35460.81");
                    table.Cell().Element(CellStyleWithoutBorders).Text("7.59");
                    table.Cell().Element(CellStyleWithoutBorders).Text("0.19");

                    table.Cell().Element(CellStyleWithoutBorders).Text("12");
                    table.Cell().Element(CellStyleWithoutBorders).Text("TENSIONER BELT V");
                    table.Cell().Element(CellStyleWithoutBorders).Text("3.37Kg");
                    table.Cell().Element(CellStyleWithoutBorders).Text("02 PCS");
                    table.Cell().Element(CellStyleWithoutBorders).Text("12932.00");
                    table.Cell().Element(CellStyleWithoutBorders).Text("1424.10");
                    table.Cell().Element(CellStyleWithoutBorders).Text("20.52");
                    table.Cell().Element(CellStyleWithoutBorders).Text("1369.13");
                    table.Cell().Element(CellStyleWithoutBorders).Text("15745.75");
                    table.Cell().Element(CellStyleWithoutBorders).Text("3.37");
                    table.Cell().Element(CellStyleWithoutBorders).Text("0.08");

                    // Totals for ITEM 004 with top and bottom borders
                    table.Cell().Element(CellStyleWithBorders).Text("");
                    table.Cell().Element(CellStyleWithBorders).Text("");
                    table.Cell().Element(CellStyleWithBorders).Text("10.96Kg");
                    table.Cell().Element(CellStyleWithBorders).Text("6.00");
                    table.Cell().Element(CellStyleWithBorders).Text("42056.00");
                    table.Cell().Element(CellStyleWithBorders).Text("4631.30");
                    table.Cell().Element(CellStyleWithBorders).Text("66.73");
                    table.Cell().Element(CellStyleWithBorders).Text("4452.53");
                    table.Cell().Element(CellStyleWithBorders).Text("51206.56");
                    table.Cell().Element(CellStyleWithBorders).Text("10.96");
                    table.Cell().Element(CellStyleWithBorders).Text("0.27");

                    // ITEM 005
                    table.Cell().Element(CellStyleWithoutBorders).Text("ITEM: 005").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("84841000").Bold();
                    for (int i = 0; i < 9; i++) table.Cell().Element(CellStyleWithoutBorders).Text("");

                    // Header row for ITEM 005
                    table.Cell().Element(CellStyleWithoutBorders).Text("No").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Description").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Qty 1").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("MU Qty2").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("FOB").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Freight").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Ins").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Others").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("CIF").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Nett").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Pkgs").Bold();

                    // Data row for ITEM 005
                    table.Cell().Element(CellStyleWithoutBorders).Text("8");
                    table.Cell().Element(CellStyleWithoutBorders).Text("GASKET");
                    table.Cell().Element(CellStyleWithoutBorders).Text("1.25Kg");
                    table.Cell().Element(CellStyleWithoutBorders).Text("200 PCS");
                    table.Cell().Element(CellStyleWithoutBorders).Text("4800.00");
                    table.Cell().Element(CellStyleWithoutBorders).Text("528.59");
                    table.Cell().Element(CellStyleWithoutBorders).Text("7.62");
                    table.Cell().Element(CellStyleWithoutBorders).Text("508.18");
                    table.Cell().Element(CellStyleWithoutBorders).Text("5844.39");
                    table.Cell().Element(CellStyleWithoutBorders).Text("1.25");
                    table.Cell().Element(CellStyleWithoutBorders).Text("0.03");

                    // Totals for ITEM 005 with top and bottom borders
                    table.Cell().Element(CellStyleWithBorders).Text(" ");
                    table.Cell().Element(CellStyleWithBorders).Text(" ");
                    table.Cell().Element(CellStyleWithBorders).Text("1.25Kg");
                    table.Cell().Element(CellStyleWithBorders).Text("200 PCS");
                    table.Cell().Element(CellStyleWithBorders).Text("4800.00");
                    table.Cell().Element(CellStyleWithBorders).Text("528.59");
                    table.Cell().Element(CellStyleWithBorders).Text("7.62");
                    table.Cell().Element(CellStyleWithBorders).Text("508.18");
                    table.Cell().Element(CellStyleWithBorders).Text("5844.39");
                    table.Cell().Element(CellStyleWithBorders).Text("1.25");
                    table.Cell().Element(CellStyleWithBorders).Text("0.03");


                    // ITEM 006
                    table.Cell().Element(CellStyleWithoutBorders).Text("ITEM: 006").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("8708990").Bold();
                    for (int i = 0; i < 9; i++) table.Cell().Element(CellStyleWithoutBorders).Text("");

                    // Define the headers similar to your example
                    table.Cell().Element(CellStyleWithoutBorders).Text("No").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Description").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Qty 1").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("MU Qty2").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("FOB").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Freight").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Ins").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Others").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("CIF").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Nett").Bold();
                    table.Cell().Element(CellStyleWithoutBorders).Text("Pkgs").Bold();

                    // Continue adding items as needed, just like the example
                    table.Cell().Element(CellStyleWithoutBorders).Text("9");
                    table.Cell().Element(CellStyleWithoutBorders).Text("JOINT (L) BALL-OUT JOINT (L) BALL-OUT");
                    table.Cell().Element(CellStyleWithoutBorders).Text("1.30Kg");
                    table.Cell().Element(CellStyleWithoutBorders).Text("04 PCS");
                    table.Cell().Element(CellStyleWithoutBorders).Text("4972.00");
                    table.Cell().Element(CellStyleWithoutBorders).Text("547.53");
                    table.Cell().Element(CellStyleWithoutBorders).Text("7.89");
                    table.Cell().Element(CellStyleWithoutBorders).Text("526.39");
                    table.Cell().Element(CellStyleWithoutBorders).Text("6053.81");
                    table.Cell().Element(CellStyleWithoutBorders).Text("1.30");
                    table.Cell().Element(CellStyleWithoutBorders).Text("0.03");

                    // Totals row with borders
                    table.Cell().Element(CellStyleWithBorders).Text("");
                    table.Cell().Element(CellStyleWithBorders).Text("");
                    table.Cell().Element(CellStyleWithBorders).Text("1.30Kg");
                    table.Cell().Element(CellStyleWithBorders).Text("04 PCS");
                    table.Cell().Element(CellStyleWithBorders).Text("4972.00");
                    table.Cell().Element(CellStyleWithBorders).Text("547.53");
                    table.Cell().Element(CellStyleWithBorders).Text("7.89");
                    table.Cell().Element(CellStyleWithBorders).Text("526.39");
                    table.Cell().Element(CellStyleWithBorders).Text("6053.81");
                    table.Cell().Element(CellStyleWithBorders).Text("1.30");
                    table.Cell().Element(CellStyleWithBorders).Text("0.03");

                });

            });

        // Footer with page number
        page.Footer()
            .AlignCenter()
            .Text(x =>
            {
                x.Span("Page ");
                x.CurrentPageNumber();
            });
    });
})
.ShowInPreviewer();


// Helper method to style the table cells without borders
IContainer CellStyleWithoutBorders(IContainer container)
{
    return container.PaddingVertical(2).AlignCenter(); // Removed borders, keeping vertical padding and alignment
}

IContainer CellStyleWithBorders(IContainer container)
{
    return container
        .BorderTop(1) 
        .BorderBottom(1)
        .PaddingVertical(2)
        .AlignCenter();
}
