package excel

// GetSheetMeta はシートのタブ色、デフォルト列幅、デフォルト行高を返す
func (f *File) GetSheetMeta(sheet string) (tabColor string, defaultWidth, defaultHeight float64, err error) {
	props, err := f.File.GetSheetProps(sheet)
	if err != nil {
		return "", 0, 0, err
	}

	// タブ色
	if props.TabColorRGB != nil && *props.TabColorRGB != "" {
		tabColor = normalizeHexColor(*props.TabColorRGB)
	} else if props.TabColorTheme != nil {
		tint := 0.0
		if props.TabColorTint != nil {
			tint = *props.TabColorTint
		}
		tabColor = ResolveColor("", props.TabColorTheme, tint, f.File)
	}

	if props.DefaultColWidth != nil {
		defaultWidth = *props.DefaultColWidth
	}
	if props.DefaultRowHeight != nil {
		defaultHeight = *props.DefaultRowHeight
	}

	return tabColor, defaultWidth, defaultHeight, nil
}
