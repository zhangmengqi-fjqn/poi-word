/*
 * Copyright 2014-2015 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.melon.word.utils;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.exception.RenderException;
import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import com.deepoove.poi.template.run.RunTemplate;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlToken;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;

import java.io.ByteArrayInputStream;
import java.io.FileInputStream;
import java.io.InputStream;

public class PictureRenderPolicyCustomer extends AbstractRenderPolicy<PictureRenderData> {

    @Override
    protected boolean validate(PictureRenderData data) {
        return (null != data.getData() || null != data.getPath());
    }

    @Override
    public void doRender(RunTemplate runTemplate, PictureRenderData picture, XWPFTemplate template)
            throws Exception {
        XWPFRun run = runTemplate.getRun();
        Helper.renderPicture(run, picture);
    }

    @Override
    protected void afterRender(RenderContext context) {
        clearPlaceholder(context, false);
    }

    @Override
    protected void doRenderException(RunTemplate runTemplate, PictureRenderData data, Exception e) {
        logger.info("Render picture " + runTemplate + " error: {}", e.getMessage());
        runTemplate.getRun().setText(data.getAltMeta(), 0);
    }

    public static class Helper {
        public static final int EMU = 9525;

        public static void renderPicture(XWPFRun run, PictureRenderData picture) throws Exception {
            InputStream ins = null == picture.getData() ? new FileInputStream(picture.getPath())
                    : new ByteArrayInputStream(picture.getData());

//            run.addPicture(ins, suggestFileType, "Generated", picture.getWidth() * EMU,
//                    picture.getHeight() * EMU);
            XWPFDocument document = run.getDocument();
            String blipId = document.addPictureData(ins, XWPFDocument.PICTURE_TYPE_PNG);
            int id = document.getNextPicNameNumber(XWPFDocument.PICTURE_TYPE_PNG);
            int width = picture.getWidth();
            int height = picture.getHeight();


            CTInline inline = run.getCTR().addNewDrawing().addNewInline();

            String picXml = "" +
                    "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                    "   <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                    "      <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                    "         <pic:nvPicPr>" +
                    "            <pic:cNvPr id=\"" + id + "\" name=\"Generated\"/>" +
                    "            <pic:cNvPicPr/>" +
                    "         </pic:nvPicPr>" +
                    "         <pic:blipFill>" +
                    "            <a:blip r:embed=\"" + blipId + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>" +
                    "            <a:stretch>" +
                    "               <a:fillRect/>" +
                    "            </a:stretch>" +
                    "         </pic:blipFill>" +
                    "         <pic:spPr>" +
                    "            <a:xfrm>" +
                    "               <a:off x=\"0\" y=\"0\"/>" +
                    "               <a:ext cx=\"" + width + "\" cy=\"" + height + "\"/>" +
                    "            </a:xfrm>" +
                    "            <a:prstGeom prst=\"rect\">" +
                    "               <a:avLst/>" +
                    "            </a:prstGeom>" +
                    "         </pic:spPr>" +
                    "      </pic:pic>" +
                    "   </a:graphicData>" +
                    "</a:graphic>";

            //CTGraphicalObjectData graphicData = inline.addNewGraphic().addNewGraphicData();
            XmlToken xmlToken = null;
            try {
                xmlToken = XmlToken.Factory.parse(picXml);
            } catch (XmlException xe) {
                xe.printStackTrace();
            }
            inline.set(xmlToken);
            //graphicData.set(xmlToken);

            inline.setDistT(0);
            inline.setDistB(0);
            inline.setDistL(0);
            inline.setDistR(0);

            CTPositiveSize2D extent = inline.addNewExtent();
            extent.setCx(width);
            extent.setCy(height);

            CTNonVisualDrawingProps docPr = inline.addNewDocPr();
            docPr.setId(id);
            docPr.setName("Picture " + id);
            docPr.setDescr("Generated");

        }
    }
}
