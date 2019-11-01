package com.melon.word.exceptions;

/**
 * @author zhaokai
 * @date 2019-11-01
 */
public class WordException extends RuntimeException {

    public WordException() {
        super();
    }

    public WordException(String message) {
        super(message);
    }

    public WordException(String message, Throwable cause) {
        super(message, cause);
    }

    public WordException(Throwable cause) {
        super(cause);
    }

    protected WordException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
