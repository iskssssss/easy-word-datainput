package test.model;

import lombok.Data;

import java.io.Serializable;

@Data
public class TaskInfoTagsVo implements Serializable {

    private String code;

    private String name;

    private Integer summarize;

    private String rate;


}
