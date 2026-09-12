interface PictureOptions {
  width: number
  height: number
  altTextTitle?: string
  altTextDescription?: string
}

export function createDiagramPictureOoxml(base64: string, options: PictureOptions): string {
  const cx = Math.round(options.width * 12700)
  const cy = Math.round(options.height * 12700)
  if (!Number.isSafeInteger(cx) || !Number.isSafeInteger(cy) || cx <= 0 || cy <= 0) {
    throw new Error('Diagram dimensions must be positive finite Word picture dimensions.')
  }

  // Import the bitmap and its complete drawing geometry together. Word's base64
  // replacement API can paint a higher-resolution bitmap outside its old frame.
  const document = new DOMParser().parseFromString(
    `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
      <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
        <pkg:xmlData>
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
          </Relationships>
        </pkg:xmlData>
      </pkg:part>
      <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
        <pkg:xmlData>
          <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">
            <w:body><w:p><w:r><w:drawing>
              <wp:inline distT="0" distB="0" distL="0" distR="0">
                <wp:extent cx="${cx}" cy="${cy}"/>
                <wp:effectExtent l="0" t="0" r="0" b="0"/>
                <wp:docPr id="1" name="Mermaid diagram"/>
                <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
                <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <pic:pic>
                    <pic:nvPicPr><pic:cNvPr id="0" name="Mermaid diagram.png"/><pic:cNvPicPr/></pic:nvPicPr>
                    <pic:blipFill>
                      <a:blip r:embed="rIdImage"><a:extLst>
                        <a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"><a14:useLocalDpi val="0"/></a:ext>
                      </a:extLst></a:blip>
                      <a:stretch><a:fillRect/></a:stretch>
                    </pic:blipFill>
                    <pic:spPr>
                      <a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
                      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                    </pic:spPr>
                  </pic:pic>
                </a:graphicData></a:graphic>
              </wp:inline>
            </w:drawing></w:r></w:p></w:body>
          </w:document>
        </pkg:xmlData>
      </pkg:part>
      <pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
        <pkg:xmlData>
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/diagram.png"/>
          </Relationships>
        </pkg:xmlData>
      </pkg:part>
      <pkg:part pkg:name="/word/media/diagram.png" pkg:contentType="image/png">
        <pkg:binaryData/>
      </pkg:part>
    </pkg:package>`,
    'application/xml',
  )
  const properties = document.getElementsByTagNameNS(
    'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing', 'docPr',
  )[0]
  properties.setAttribute('title', options.altTextTitle ?? 'Mermaid diagram')
  properties.setAttribute('descr', options.altTextDescription ?? 'Diagram created with Mermaid Office.')
  document.getElementsByTagNameNS(
    'http://schemas.microsoft.com/office/2006/xmlPackage', 'binaryData',
  )[0].textContent = base64
  return new XMLSerializer().serializeToString(document)
}
