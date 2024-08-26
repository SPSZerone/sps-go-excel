package excel

type Option func(*Options)
type Flag uint32

// Flags to OpenFile OpenExcel OpenExcelCustom NewFile NewExcel NewExcelCustom wrapping those of the underlying system.
// Exactly one of OReadOnly, OWriteOnly, or OReadWrite must be specified.
// The remaining values may be or'ed in to control behavior.
const (
	OReadOnly  Flag = 1 << iota // open the file read-only.
	OWriteOnly                  // open the file write-only.
	OReadWrite                  // open the file read-write.

	OAppend // append data to the file when writing.
	OCreate // create a new file if none exists.
	OExist  // used with OCreate, file must not exist.
)

func newOptions(opts ...Option) *Options {
	options := &Options{}
	for _, option := range opts {
		option(options)
	}
	return options
}

type Options struct {
	File string
	Flag Flag
}

func OptFile(value string) Option {
	return func(c *Options) {
		c.File = value
	}
}

func OptFlag(value Flag) Option {
	return func(c *Options) {
		c.Flag = value
	}
}

func (f Flag) IsCreate() bool {
	return f.Check(OCreate)
}

func (f Flag) IsExist() bool {
	return f.Check(OExist)
}

func (f Flag) IsWritable() bool {
	return f.Check(OWriteOnly | OReadWrite)
}

func (f Flag) Check(flag Flag) bool {
	return (f & flag) != 0
}
