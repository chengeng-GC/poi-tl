package com.deepoove.poi.util;

import org.apache.commons.io.input.UnsynchronizedByteArrayInputStream;
import org.apache.commons.io.output.UnsynchronizedByteArrayOutputStream;
import org.apache.poi.poifs.filesystem.*;
import org.apache.poi.util.*;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;

/**
 * 类名称: TestOLE
 * 类描述:
 * 创建人: 陈庚
 * 创建时间: 2024/11/27 16:28
 * 修改人: 陈庚
 * 修改时间: 2024/11/27 16:28
 * copyright (c) 善算科技有限公司
 * 修改备注:
 */
public class TestOLE {
    public static final String OLE10_NATIVE = "\u0001TestOLE";
    private static final Charset GBK;
    private static final Charset ISO;
    private static final int DEFAULT_MAX_RECORD_LENGTH = 100000000;
    private static int MAX_RECORD_LENGTH;
    private static final int DEFAULT_MAX_STRING_LENGTH = 1024;
    private static int MAX_STRING_LENGTH;
    private static final byte[] OLE_MARKER_BYTES;
    private static final String OLE_MARKER_NAME = "\u0001Ole";
    private int totalSize;
    private short flags1 = 2;
    private String label;
    private String fileName;
    private short flags2;
    private short unknown1 = 3;
    private String command;
    private byte[] dataBuffer;
    private String command2;
    private String label2;
    private String fileName2;
    private TestOLE.EncodingMode mode;

    public static TestOLE createFromEmbeddedOleObject(POIFSFileSystem poifs) throws IOException, Ole10NativeException {
        return createFromEmbeddedOleObject(poifs.getRoot());
    }

    public static TestOLE createFromEmbeddedOleObject(DirectoryNode directory) throws IOException, Ole10NativeException {
        DocumentEntry nativeEntry = (DocumentEntry)directory.getEntry("\u0001TestOLE");
        DocumentInputStream dis = directory.createDocumentInputStream(nativeEntry);
        Throwable var3 = null;

        TestOLE var5;
        try {
            byte[] data = IOUtils.toByteArray(dis, nativeEntry.getSize(), MAX_RECORD_LENGTH);
            var5 = new TestOLE(data, 0);
        } catch (Throwable var14) {
            var3 = var14;
            throw var14;
        } finally {
            if (dis != null) {
                if (var3 != null) {
                    try {
                        dis.close();
                    } catch (Throwable var13) {
                        var3.addSuppressed(var13);
                    }
                } else {
                    dis.close();
                }
            }

        }

        return var5;
    }

    public static void setMaxRecordLength(int length) {
        MAX_RECORD_LENGTH = length;
    }

    public static int getMaxRecordLength() {
        return MAX_RECORD_LENGTH;
    }

    public static void setMaxStringLength(int length) {
        MAX_STRING_LENGTH = length;
    }

    public static int getMaxStringLength() {
        return MAX_STRING_LENGTH;
    }

    public TestOLE(String label, String filename, String command, byte[] data) {
        this.setLabel(label);
        this.setFileName(filename);
        this.setCommand(command);
        this.command2 = command;
        this.setDataBuffer(data);
        this.mode = TestOLE.EncodingMode.parsed;
    }

    public TestOLE(byte[] data, int offset) throws Ole10NativeException {
        LittleEndianByteArrayInputStream leis = new LittleEndianByteArrayInputStream(data, offset);
        this.totalSize = leis.readInt();
        leis.limit(this.totalSize + 4);
        leis.mark(0);

        try {
            this.flags1 = leis.readShort();
            if (this.flags1 == 2) {
                leis.mark(0);
                boolean validFileName = !Character.isISOControl(leis.readByte());
                leis.reset();
                if (validFileName) {
                    this.readParsed(leis);
                } else {
                    this.readCompact(leis);
                }
            } else {
                leis.reset();
                this.readUnparsed(leis);
            }

        } catch (IOException var5) {
            throw new Ole10NativeException("Invalid TestOLE", var5);
        }
    }

    private void readParsed(LittleEndianByteArrayInputStream leis) throws Ole10NativeException, IOException {
        this.mode = TestOLE.EncodingMode.parsed;
        this.label = readAsciiZ(leis);
        this.fileName = readAsciiZ(leis);
        this.flags2 = leis.readShort();
        this.unknown1 = leis.readShort();
        this.command = readAsciiLen(leis);
        this.dataBuffer = IOUtils.toByteArray(leis, leis.readInt(), MAX_RECORD_LENGTH);
        leis.mark(0);
        short lowSize = leis.readShort();
        if (lowSize != 0) {
            leis.reset();
            this.command2 = readUtf16(leis);
            this.label2 = readUtf16(leis);
            this.fileName2 = readUtf16(leis);
        }

    }

    private void readCompact(LittleEndianByteArrayInputStream leis) throws IOException {
        this.mode = TestOLE.EncodingMode.compact;
        this.dataBuffer = IOUtils.toByteArray(leis, this.totalSize - 2, MAX_RECORD_LENGTH);
    }

    private void readUnparsed(LittleEndianByteArrayInputStream leis) throws IOException {
        this.mode = TestOLE.EncodingMode.unparsed;
        this.dataBuffer = IOUtils.toByteArray(leis, this.totalSize, MAX_RECORD_LENGTH);
    }

    public static void createOleMarkerEntry(DirectoryEntry parent) throws IOException {
        if (!parent.hasEntry("\u0001Ole")) {
            parent.createDocument("\u0001Ole", new UnsynchronizedByteArrayInputStream(OLE_MARKER_BYTES));
        }

    }

    public static void createOleMarkerEntry(POIFSFileSystem poifs) throws IOException {
        createOleMarkerEntry((DirectoryEntry)poifs.getRoot());
    }

    private static String readAsciiZ(LittleEndianInput is) throws Ole10NativeException {
        byte[] buf = new byte[MAX_STRING_LENGTH];

        for(int i = 0; i < buf.length; ++i) {
            if ((buf[i] = is.readByte()) == 0) {
                return StringUtil.getFromCompressedUnicode(buf, 0, i);
            }
        }

        throw new Ole10NativeException("AsciiZ string was not null terminated after " + MAX_STRING_LENGTH + " bytes - Exiting.");
    }

    private static String readAsciiLen(LittleEndianByteArrayInputStream leis) throws IOException {
        int size = leis.readInt();
        byte[] buf = IOUtils.toByteArray(leis, size, MAX_STRING_LENGTH);
        return buf.length == 0 ? "" : StringUtil.getFromCompressedUnicode(buf, 0, size - 1);
    }

    private static String readUtf16(LittleEndianByteArrayInputStream leis) throws IOException {
        int size = leis.readInt();
        byte[] buf = IOUtils.toByteArray(leis, size * 2, MAX_STRING_LENGTH);
        return StringUtil.getFromUnicodeLE(buf, 0, size);
    }

    public int getTotalSize() {
        return this.totalSize;
    }

    public short getFlags1() {
        return this.flags1;
    }

    public String getLabel() {
        return this.label;
    }

    public String getFileName() {
        return this.fileName;
    }

    public short getFlags2() {
        return this.flags2;
    }

    public short getUnknown1() {
        return this.unknown1;
    }

    public String getCommand() {
        return this.command;
    }

    public int getDataSize() {
        return this.dataBuffer.length;
    }

    public byte[] getDataBuffer() {
        return this.dataBuffer;
    }

    public void writeOut(OutputStream out) throws IOException {
        LittleEndianOutputStream leosOut = new LittleEndianOutputStream(out);
        switch (this.mode) {
            case parsed:
                UnsynchronizedByteArrayOutputStream bos = new UnsynchronizedByteArrayOutputStream();
                LittleEndianOutputStream leos = new LittleEndianOutputStream(bos);
                Throwable var5 = null;

                try {
                    leos.writeShort(this.getFlags1());
                    leos.write(this.getLabel().getBytes(GBK));
                    leos.write(0);
                    leos.write(this.getFileName().getBytes(GBK));
                    leos.write(0);
                    leos.writeShort(this.getFlags2());
                    leos.writeShort(this.getUnknown1());
                    leos.writeInt(this.getCommand().getBytes(GBK).length + 1);
                    leos.write(this.getCommand().getBytes(GBK));
                    leos.write(0);
                    leos.writeInt(this.getDataSize());
                    leos.write(this.getDataBuffer());
                    if (this.command2 != null && this.label2 != null && this.fileName2 != null) {
                        leos.writeUInt((long)this.command2.length());
                        leos.write(StringUtil.getToUnicodeLE(this.command2));
                        leos.writeUInt((long)this.label2.length());
                        leos.write(StringUtil.getToUnicodeLE(this.label2));
                        leos.writeUInt((long)this.fileName2.length());
                        leos.write(StringUtil.getToUnicodeLE(this.fileName2));
                    } else {
                        leos.writeShort(0);
                    }
                } catch (Throwable var14) {
                    var5 = var14;
                    throw var14;
                } finally {
                    if (var5 != null) {
                        try {
                            leos.close();
                        } catch (Throwable var13) {
                            var5.addSuppressed(var13);
                        }
                    } else {
                        leos.close();
                    }

                }

                leosOut.writeInt(bos.size());
                bos.writeTo(out);
                break;
            case compact:
                leosOut.writeInt(this.getDataSize() + 2);
                leosOut.writeShort(this.getFlags1());
                out.write(this.getDataBuffer());
                break;
            case unparsed:
            default:
                leosOut.writeInt(this.getDataSize());
                out.write(this.getDataBuffer());
        }

    }

    public void setFlags1(short flags1) {
        this.flags1 = flags1;
    }

    public void setFlags2(short flags2) {
        this.flags2 = flags2;
    }

    public void setLabel(String label) {
        this.label = label;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public void setCommand(String command) {
        this.command = command;
    }

    public void setUnknown1(short unknown1) {
        this.unknown1 = unknown1;
    }

    public void setDataBuffer(byte[] dataBuffer) {
        this.dataBuffer = (byte[])dataBuffer.clone();
    }

    public String getCommand2() {
        return this.command2;
    }

    public void setCommand2(String command2) {
        this.command2 = command2;
    }

    public String getLabel2() {
        return this.label2;
    }

    public void setLabel2(String label2) {
        this.label2 = label2;
    }

    public String getFileName2() {
        return this.fileName2;
    }

    public void setFileName2(String fileName2) {
        this.fileName2 = fileName2;
    }

    static {
        GBK = Charset.forName("GBK");
        ISO = StandardCharsets.ISO_8859_1;
        MAX_RECORD_LENGTH = 100000000;
        MAX_STRING_LENGTH = 1024;
        OLE_MARKER_BYTES = new byte[]{1, 0, 0, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0};
    }

    private static enum EncodingMode {
        parsed,
        unparsed,
        compact;

        private EncodingMode() {
        }
    }
}
