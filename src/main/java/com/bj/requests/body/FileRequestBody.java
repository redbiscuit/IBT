package com.bj.requests.body;

import com.bj.requests.HttpHeaders;
import com.bj.requests.utils.InputOutputs;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.Charset;
import java.nio.file.Files;

/**
 * @author Liu Dong
 */
class FileRequestBody extends RequestBody<File> {
    FileRequestBody(File body) {
        super(body, getFileContentType(body), false);
    }

    private static String getFileContentType(File body) {
        String contentType;
        try {
            contentType = Files.probeContentType(body.toPath());
        } catch (IOException e) {
            contentType = null;
        }
        if (contentType == null) {
            contentType = HttpHeaders.CONTENT_TYPE_BINARY;
        }
        return contentType;
    }

    @Override
    public void writeBody(OutputStream os, Charset charset) throws IOException {
        InputOutputs.copy(new FileInputStream(getBody()), os);
    }
}
