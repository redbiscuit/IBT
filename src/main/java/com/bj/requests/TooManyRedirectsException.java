package com.bj.requests;

/**
 * Thrown when redirect too many times.
 */
public class TooManyRedirectsException extends RequestsException {
    public TooManyRedirectsException(int count) {
        super("Redirect too many times: " + count);
    }
}
