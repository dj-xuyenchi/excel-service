import lombok.Getter;
import lombok.Setter;

import java.math.BigDecimal;
import java.util.Date;

@Getter
@Setter
public class TestClass {
    private String name;
    private Date createdAt;
    private BigDecimal price;
}
