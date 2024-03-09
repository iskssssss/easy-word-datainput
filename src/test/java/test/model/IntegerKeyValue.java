package test.model;

import lombok.Data;
import lombok.ToString;

import java.io.Serializable;

/**
 * 键值对 数值类型
 * @author lyq
 * @date 2023-7-18 16:40
 */
@Data
@ToString
public class IntegerKeyValue implements Serializable {
    static final long serialVersionUID = 42L;

    private String key;
    private Integer value;
}
