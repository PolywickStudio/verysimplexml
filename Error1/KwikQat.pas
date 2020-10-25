unit KwikQat;
interface

procedure RepairQAT(sPath: string);

implementation
uses
  Windows,
  dialogs,
  Xml.VerySimple,
  Classes,
  SysUtils;

 function FileSize(const aFilename: String): Int64;
  var
    info: TWin32FileAttributeData;
  begin
    result := -1;

    if NOT GetFileAttributesEx(PWideChar(aFileName), GetFileExInfoStandard, @info) then
      EXIT;

    result := Int64(info.nFileSizeLow) or Int64(info.nFileSizeHigh shl 32);
  end;

procedure RepairQAT(sPath: string);
var
  xml: TXmlVerySimple;
  cui: TXMLNode;
  rib: TXMLNode;
  qat: TXMLNode;
  tab: TXMLNode;
  shb: TXMLNode;
  iNo, iCount: integer;
  sidq: string;
  //
  shc: TXMLNode;
  iFirst: integer;
  ts: TStringlist;
  bOK: boolean;
  sPathOut: string;
begin
  bOK := false;
  if FileExists(sPath) then begin
    if (FileSize(sPath) >= 200) then
    begin
      bOK := true;
    end;
  end;

  sPathOut := ChangeFileExt(sPath, '.officeui');

  if bOK then
  begin
    xml := TXmlVerySimple.Create;
    xml.LoadFromFile(sPath);
    //
    cui := xml.ChildNodes.Get(0);
    cui.Attributes['xmlns:x2'] := 'KwikDocs.CoKwikDocs';

    iFirst := -1;
    //
    tab := nil;
    shb := nil;
    qat := nil;

    if cui.ChildNodes.HasNode('mso:ribbon') then begin
      rib := cui.ChildNodes.FindNode('mso:ribbon');
    end else begin
      rib := TXmlNode.Create();
      cui.ChildNodes.Add(rib);
    end;

    if assigned(rib) then begin
      if rib.ChildNodes.HasNode('mso:qat') then begin
        qat := rib.ChildNodes.FindNode('mso:qat');
      end else begin
        qat := TXmlNode.Create();
        rib.ChildNodes.Add(qat);
      end;
      if rib.ChildNodes.HasNode('mso:tabs') then begin
        tab := rib.ChildNodes.FindNode('mso:tabs');
      end else begin
        tab := TXmlNode.Create();
        tab.NodeName := 'mso:tabs';
        rib.ChildNodes.Add(tab);
      end;
    end;

    if assigned(qat) then begin
      if qat.ChildNodes.HasNode('mso:sharedControls') then begin
        shb := qat.ChildNodes.FindNode('mso:sharedControls');
      end else begin
        shb := TXmlNode.Create();
        tab.NodeName := 'mso:tabs';
        qat.ChildNodes.Add(tab);
      end;
    end;


    if assigned(shb) then begin
      iCount := shb.ChildNodes.Count  - 1;
      for iNo := iCount downto 0 do begin
        sidq := shb.ChildNodes.Get(iNo).Attributes['idQ'];
        if (sidq = 'mso:StyleGalleryClassic') or
          (sidq.Contains('kdBtnBlankDoc')) then begin
          shb.ChildNodes.Delete(iNo);
          if iFirst = -1 then begin
            iFirst := iNo - 1;
          end;
        end;
      end;
    end;

    if assigned(tab) then begin
      iCount := tab.ChildNodes.Count  - 1;
      for iNo := iCount downto 0 do begin
        sidq := tab.ChildNodes.Get(iNo).Attributes['idQ'];
        if sidq.Contains('kdNumbering') then begin
          tab.ChildNodes.Delete(iNo);
        end;
        if sidq.Contains('kdCUQ') then begin
          tab.ChildNodes.Delete(iNo);
        end;
      end;
    end;

    shc := TXmlNode.Create();
    shc.Attributes['idQ'] := 'mso:StyleGalleryClassic';
    shc.Attributes['visible'] := 'true';
    shc.Name := 'mso:control';
    if iFirst = -1 then begin
      shb.ChildNodes.Add(shc);
    end else begin
      shb.ChildNodes.Insert(iFirst, shc);
    end;

    //====================================================================================
    shc := TXmlNode.Create();
    shc.Attributes['idQ'] := 'x2:kdBtnBlankDoc';
    shc.Attributes['visible'] := 'true';
    shc.Name := 'mso:control';
    if iFirst = -1 then begin
      shb.ChildNodes.Add(shc);
    end else begin
      shb.ChildNodes.Insert(iFirst, shc);
    end;
    //
    //====================================================================================
    shc := TXmlNode.Create();
    shc.Attributes['idQ'] := 'x2:kdNumbering';
    shc.Name := 'mso:tab';
    tab.ChildNodes.Add(shc);
    //

    shc := TXmlNode.Create();
    shc.Attributes['idQ'] := 'x2:kdCUQ';
    shc.Name := 'mso:tab';
    tab.ChildNodes.Add(shc);

    xml.SaveToFile(sPathOut);

    xml.Free;
  end else begin
    // save a default file.
    ts := tstringlist.create;
    ts.add(
      '<mso:customUI ' +
      'xmlns:x2="KwikDocs.CoKwikDocs" xmlns:mso="http://schemas.microsoft.com/office/2009/07/customui">' +
      '<mso:ribbon>' +
      '<mso:qat>' +
      '<mso:sharedControls>' +
      '<mso:control idQ="x2:kdBtnBlankDoc" visible="true"/>' +
      '<mso:control idQ="mso:StyleGalleryClassic" visible="true"/>' +
      '</mso:sharedControls>' +
      '</mso:qat>' +
      '<mso:tabs>' +
      '<mso:tab idQ="mso:TabDrawInk" visible="false"/>' +
      '<mso:tab idQ="x2:kdNumbering"/>' +
      '<mso:tab idQ="x2:kdCUQ"/>' +
      '</mso:tabs>' +
      '</mso:ribbon>' +
      '</mso:customUI>');
    //
    ts.SaveToFile(sPathOut);
    ts.Free;
  end;
end;

end.

