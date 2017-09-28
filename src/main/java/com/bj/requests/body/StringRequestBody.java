package com.bj.requests.body;

import com.bj.requests.HttpHeaders;

import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.nio.charset.Charset;

/**
 * @author Liu Dong
 */
class StringRequestBody extends RequestBody<String> {
    StringRequestBody(String body) {
        super(body, HttpHeaders.CONTENT_TYPE_TEXT, true);
    }

    @Override public void writeBody(OutputStream os, Charset charset) throws IOException {
        try (Writer writer = new OutputStreamWriter(os, charset)) {
            writer.write(getBody());
        }
    }
}
