package poi;

import java.util.Map;

public interface objectMapper {
	<T> T parseObject(Map map);
}
