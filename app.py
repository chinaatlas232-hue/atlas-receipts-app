# معالجة قراءة وتحويل الأعمدة الرقمية وتلافي الأخطاء
    df["الوزن"] = pd.to_numeric(df.get("الوزن", 0), errors="coerce").fillna(0.0)
    df["عدد الطرود"] = pd.to_numeric(
        df.get("عدد الطرود", 0), errors="coerce"
    ).fillna(0)

    # معالجة عمود الحجم بناءً على الاختيار اليدوي أو البحث الذكي عن القيم الرقمية الحقيقية
    if manual_volume_col != "تلقائي" and manual_volume_col in df.columns:
      df["الحجم"] = pd.to_numeric(df[manual_volume_col], errors="coerce").fillna(
          0.0
      )
    else:
      found_vol_col = None
      for col in df.columns:
        clean_c = col.lower().strip()
        if any(
            k in clean_c
            for k in [
                "cbm",
                "vol",
                "volume",
                "حجم",
                "فاليوم",
                "فاليم",
                "حج",
                "الابعاد",
            ]
        ):
          # التحقق من أن العمود يحتوي على أرقام وليس رموز شحنات مثل RA6003
          converted_series = pd.to_numeric(df[col], errors="coerce")
          if converted_series.notna().sum() > 0:
            found_vol_col = col
            break
      if found_vol_col:
        df["الحجم"] = pd.to_numeric(df[found_vol_col], errors="coerce").fillna(
            0.0
        )
      else:
        df["الحجم"] = 0.0
