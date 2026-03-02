package me.muban.docx;

/**
 * Base exception for all muban-docx library errors.
 *
 * <p>This unchecked exception is the root of the muban-docx exception hierarchy.
 * Callers can catch this single type to handle all library-specific errors.
 */
public class MubanDocxException extends RuntimeException {

    private final String errorCode;

    /**
     * Create an exception with a message.
     *
     * @param message human-readable error description
     */
    public MubanDocxException(String message) {
        this("MUBAN_DOCX_ERROR", message, null);
    }

    /**
     * Create an exception with a message and cause.
     *
     * @param message human-readable error description
     * @param cause   the underlying cause
     */
    public MubanDocxException(String message, Throwable cause) {
        this("MUBAN_DOCX_ERROR", message, cause);
    }

    /**
     * Create an exception with an error code, message, and optional cause.
     *
     * @param errorCode machine-readable error code (e.g., "EXPORT_FAILED")
     * @param message   human-readable error description
     * @param cause     the underlying cause, or null
     */
    public MubanDocxException(String errorCode, String message, Throwable cause) {
        super(message, cause);
        this.errorCode = errorCode;
    }

    /**
     * @return the machine-readable error code
     */
    public String getErrorCode() {
        return errorCode;
    }
}
