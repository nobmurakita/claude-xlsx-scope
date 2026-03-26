package excel

// Truncation は切り捨て通知
type Truncation struct {
	Truncated bool   `json:"_truncated"`
	Total     int    `json:"total"`
	Output    int    `json:"output"`
	NextCell  string `json:"next_cell"`
	NextRange string `json:"next_range,omitempty"`
}
