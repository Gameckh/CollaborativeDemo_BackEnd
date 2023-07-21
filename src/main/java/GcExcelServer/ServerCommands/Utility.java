package GcExcelServer.ServerCommands;

public class Utility {
    public static int ParseARGBColorFromHtml(String color)
    {
        if (color.startsWith("#"))
        {
            String colorText = color.substring(1);

            if (colorText.length() == 3) // Convert 3 digits color to 6 digits
            {
                StringBuilder correctColorText = new StringBuilder();
                correctColorText.append(colorText.charAt(0));
                correctColorText.append(colorText.charAt(0));
                correctColorText.append(colorText.charAt(1));
                correctColorText.append(colorText.charAt(1));
                correctColorText.append(colorText.charAt(2));
                correctColorText.append(colorText.charAt(2));

                colorText = correctColorText.toString();
            }

            // 6 digits
            int argb = 0;
            RefObject<Integer> tempRef_argb = new RefObject<Integer>(argb);
            if (TryParseHelper.tryParseInt(colorText, tempRef_argb, 16)) {
                argb = tempRef_argb.argValue;
                return argb | 0xFF000000;
            }
        }

        return  0;
    }
}
