package test.model;

import lombok.Data;
import lombok.ToString;

import java.io.Serializable;

/**
 * 键值对 数值类型 带id
 * @author lyq
 * @date 2023-10-17 10:59
 */
@Data
@ToString
public class IntegerKeyValueWithId implements Serializable {
    static final long serialVersionUID = 42L;
    private String id;
    private String key;
    private Integer value;
}
