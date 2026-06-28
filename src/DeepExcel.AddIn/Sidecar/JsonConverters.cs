using System;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace DeepExcel.AddIn.Sidecar
{
    /// <summary>
    /// 将二维数组（T[,]）序列化为交错数组（T[][]），解决 System.Text.Json
    /// 不支持二维数组序列化导致的 NotSupportedException 崩溃。
    /// 应用于 object[,] 和 string[,]（RangeInfo.Values / Formulas / NumberFormats 等）。
    /// </summary>
    public class TwoDimensionalArrayConverter<T> : JsonConverter<T[,]>
    {
        public override T[,] Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            // 反序列化不常用（sidecar → C# 方向不传 2D 数组），简单支持交错数组 → 2D
            if (reader.TokenType != JsonTokenType.StartArray)
            {
                reader.Skip();
                return new T[0, 0];
            }

            var rows = new System.Collections.Generic.List<T[]>();
            while (reader.Read() && reader.TokenType != JsonTokenType.EndArray)
            {
                var row = JsonSerializer.Deserialize<T[]>(ref reader, options);
                rows.Add(row);
            }

            if (rows.Count == 0) return new T[0, 0];
            int cols = rows[0]?.Length ?? 0;
            var result = new T[rows.Count, cols];
            for (int i = 0; i < rows.Count; i++)
            {
                for (int j = 0; j < cols && j < (rows[i]?.Length ?? 0); j++)
                {
                    result[i, j] = rows[i][j];
                }
            }
            return result;
        }

        public override void Write(Utf8JsonWriter writer, T[,] value, JsonSerializerOptions options)
        {
            if (value == null)
            {
                writer.WriteNullValue();
                return;
            }

            // ★ Excel COM 返回的二维数组是 1-based（GetLowerBound 返回 1），
            // 不能用 0-based 索引访问，否则 IndexOutOfRangeException。
            // 用 GetLowerBound/GetUpperBound 安全遍历任意下界的数组。
            int rowStart = value.GetLowerBound(0);
            int rowEnd = value.GetUpperBound(0);
            int colStart = value.GetLowerBound(1);
            int colEnd = value.GetUpperBound(1);
            writer.WriteStartArray();
            for (int i = rowStart; i <= rowEnd; i++)
            {
                writer.WriteStartArray();
                for (int j = colStart; j <= colEnd; j++)
                {
                    T item;
                    try
                    {
                        item = value[i, j];
                    }
                    catch
                    {
                        // 极端情况：越界访问（理论上不应发生，防御性编程）
                        writer.WriteNullValue();
                        continue;
                    }

                    if (item == null)
                    {
                        writer.WriteNullValue();
                    }
                    else if (item is string || item is int || item is long || item is double ||
                             item is float || item is decimal || item is bool || item is DateTime)
                    {
                        // 基础类型直接写入，避免递归到 ObjectDefaultConverter
                        JsonSerializer.Serialize(writer, item, item.GetType(), options);
                    }
                    else
                    {
                        // ★ 非 base type 的元素（可能是 COM 对象如 Excel ErrorValue、嵌套 Range 等）
                        // 强制转为字符串，避免递归序列化导致 IndexOutOfRangeException 或死循环。
                        // 这是 range.Value2 返回的二维数组中元素可能不是基础类型的兜底处理。
                        string strVal;
                        try
                        {
                            strVal = item.GetType().IsValueType ? Convert.ToString(item) : item.ToString();
                        }
                        catch
                        {
                            strVal = item.GetType().Name;
                        }
                        writer.WriteStringValue(strVal);
                    }
                }
                writer.WriteEndArray();
            }
            writer.WriteEndArray();
        }
    }

    /// <summary>
    /// object[,] 专用转换器（Values 字段，元素可能是 string/double/bool/null/DateTime 等）
    /// </summary>
    public class Object2DArrayConverter : TwoDimensionalArrayConverter<object> { }

    /// <summary>
    /// string[,] 专用转换器（Formulas / NumberFormats 字段）
    /// </summary>
    public class String2DArrayConverter : TwoDimensionalArrayConverter<string> { }
}
