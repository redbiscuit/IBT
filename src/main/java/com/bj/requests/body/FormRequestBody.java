package com.bj.requests.body;

import com.bj.requests.HttpHeaders;
import com.bj.requests.URIEncoder;

import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.nio.charset.Charset;
import java.util.Collection;
import java.util.Map;

/**
 * @author Liu Dong
 */
class FormRequestBody extends RequestBody<Collection<? extends Map.Entry<String, ?>>> {
    FormRequestBody(Collection<? extends Map.Entry<String, ?>> body) {
        super(body, HttpHeaders.CONTENT_TYPE_FORM_ENCODED, true);
    }

    @Override
    public void writeBody(OutputStream os, Charset charset) throws IOException {
        String content = URIEncoder.encodeForms(URIEncoder.toStringParameters(getBody()), charset);
        try (Writer writer = new OutputStreamWriter(os, charset)) {
            writer.write(content);
        }
    }

}
