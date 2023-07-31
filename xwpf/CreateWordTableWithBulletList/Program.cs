//This sample is from the following post
//https://stackoverflow.com/questions/39510069/apache-poi-bullets-and-numbering

using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;

using (XWPFDocument document = new XWPFDocument())
{

    XWPFParagraph paragraph = document.CreateParagraph();
    XWPFRun run = paragraph.CreateRun();
    run.SetText("The table:");

    XWPFTable ltable = document.CreateTable(1, 1);

    ltable.GetCTTbl().AddNewTblGrid().AddNewGridCol().w = 5000;
    CT_TblWidth tblWidth = ltable.GetRow(0).GetCell(0).GetCTTc().AddNewTcPr().AddNewTcW();
    tblWidth.w = "5000";
    tblWidth.type = ST_TblWidth.dxa;

    ltable.GetRow(0).GetCell(0).Paragraphs[0].CreateRun().SetText("The list:");

    List<String> documentList = new List<String>(
      new String[] {
     "documentList item 1",
     "documentList item 2",
     "documentList item 3"
      });


    //your code with supplements

    CT_AbstractNum cTAbstractNum = new CT_AbstractNum();
    //Next we set the AbstractNumId. This requires care. 
    //Since we are in a new document we can start numbering from 0. 
    //But if we have an existing document, we must determine the next free number first.
    cTAbstractNum.abstractNumId = "0";

    ///* Bullet list
    cTAbstractNum.lvl = new List<CT_Lvl>();
    CT_Lvl cTLvl = new CT_Lvl();
    cTAbstractNum.lvl.Add(cTLvl);
    cTLvl.ilvl = "0"; // set indent level 0
    cTLvl.numFmt = new CT_NumFmt();
    cTLvl.numFmt.val = ST_NumberFormat.bullet;
    cTLvl.lvlText = new CT_LevelText();
    cTLvl.lvlText.val = "\u2022";
    //*/

    /* Decimal list
      CTLvl cTLvl = cTAbstractNum.addNewLvl();
      cTLvl.SetIlvl(BigInteger.valueOf(0)); // set indent level 0
      cTLvl.addNewNumFmt().SetVal(STNumberFormat.DECIMAL);
      cTLvl.addNewLvlText().SetVal("%1.");
      cTLvl.addNewStart().SetVal(BigInteger.valueOf(1));
    */

    XWPFAbstractNum abstractNum = new XWPFAbstractNum(cTAbstractNum);

    XWPFNumbering numbering = document.CreateNumbering();

    string abstractNumID = numbering.AddAbstractNum(abstractNum);

    string numID = numbering.AddNum(abstractNumID);

    foreach (String str in documentList)
    {
        XWPFTableRow lnewRow = ltable.CreateRow();
        XWPFTableCell lnewCell = lnewRow.GetCell(0);
        XWPFParagraph lnewPara = null;
        if (lnewCell.Paragraphs.Count > 0)
        {
            lnewPara = lnewCell.Paragraphs[0];
        }
        else
        {
            lnewPara = lnewCell.AddParagraph();
        }
        lnewPara.SetNumID(numID);
        XWPFRun lnewRun = lnewPara.CreateRun();
        lnewRun.SetText(str);
    }

    //your code end

    paragraph = document.CreateParagraph();

    using (FileStream sw = new FileStream("table-with-bullet.docx", FileMode.Create))
    {
        document.Write(sw);
    }
}