package org.example.libs;

import org.json.JSONArray;
import org.json.JSONObject;

public class Response {
    private boolean status;
    private String message;

    private JSONArray obj_message;
    private int code;

    public Response(boolean status, String message) {
        this.status = status;
        this.message = message;
    }

    public Response(int code, String message) {
        this.code = code;
        this.message = message;
    }

    public Response(int code, JSONArray obj_message) {
        this.code = code;
        this.obj_message = obj_message;
    }

    // Getters for status and message
    public boolean getStatus() {
        return status;
    }
    public String getMessage() {
        return message;
    }

    public void setObj_message(JSONArray obj_message) {
        this.obj_message = obj_message;
    }

    public void setCode(int code) {
        this.code = code;
    }

    //Getters for code
    public int getCode() {
        return code;
    }

    //Getter for object
    public JSONArray getResponse(){
        return obj_message;
    }

    // Setters
    public void setStatus(boolean status) {
        this.status = status;
    }

    public void setMessage(String message) {
        this.message = message;
    }

        @Override
        public String toString() {
            StringBuilder sb = new StringBuilder();
            sb.append("Response:{");

            if (status != false) {
                sb.append("status:").append(status).append(", ");
            }
            if (code != 0) {
                sb.append("code:").append(code).append(", ");
            }
            if (message != null && !message.isEmpty()) {
                sb.append("message:'").append(message).append("', ");
            }
            if (obj_message != null && obj_message.length() > 0) {
                sb.append("data:{");
                for (int i = 0; i < obj_message.length(); i++) {
                    JSONObject item = obj_message.getJSONObject(i);
                    sb.append(item.toString());
                    if (i < obj_message.length() - 1) {
                        sb.append(", ");
                    }
                }
                sb.append("}");
            }

            sb.append("}");
            return sb.toString();
        }


}
