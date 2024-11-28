/*
 * Copyright 2014-2024 Sayi
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.deepoove.poi.data;

import java.nio.file.Paths;
import java.util.UUID;
import java.util.concurrent.ThreadLocalRandom;

/**
 * attachment file
 *
 * @author Sayi
 */
public abstract class AttachmentRenderData implements RenderData {

    private static final long serialVersionUID = 1L;

    private AttachmentType fileType;
    private PictureRenderData icon;
    private String fileName;

    public abstract byte[] readAttachmentData();

    public AttachmentType getFileType() {
        if (null == fileType) setFileType(detectFileType());
        return fileType;
    }

    public void setFileType(AttachmentType fileType) {
        this.fileType = fileType;
    }

    public PictureRenderData getIcon() {
        if (null == icon) setIcon(Pictures.ofBase64(fileType.icon(), PictureType.PNG).size(64, 64).create());
        return icon;
    }

    public void setIcon(PictureRenderData icon) {
        this.icon = icon;
    }

    public String getFileName() {
        if (null == fileName){
            if (null != getFileSrc()){
                //File or Url
               setFileName(Paths.get(getFileSrc()).getFileName().toString());
            }else {
                //Byte
               String uuidRandom = UUID.randomUUID().toString().replace("-", "") + ThreadLocalRandom.current().nextInt(1024);
               setFileName(uuidRandom+fileType.ext());
            }
        }
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    protected String getFileSrc() {
        return null;
    }

    protected AttachmentType detectFileType() {
        if (null == getFileSrc()) {
            //Byte
            AttachmentType type = AttachmentType.suggestFileType(readAttachmentData());
            return type;
        }else {
            //File or Url
            AttachmentType type = AttachmentType.suggestFileType(getFileSrc());
            return type;
        }
    }


}
