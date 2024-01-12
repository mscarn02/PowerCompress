#DOWNSIZE A PPTX FILE BY COMPRESSING THE IMAGES IN IT


import os
import argparse		#used to parse command line arguments
import inspect

'''	
filename: name of the file; fname_filter: Convert images matching this filenameglob pattern;
fsize_filter: Convert images with file size larger than this limit; convert_to: Convert images to this fomat;
img_max_size: If an image is larger than this limit (in px) reduce the image to this size; quality: Save images with this quality;
optimize: Attempt to optimize the image output; fill_color: if converting images, use this color as background;
outputfn_fmt: The filename format of he generated pptx file; wait_before_zip: If true, prompt the user before zipping the files in temporry directory;
compress_type: Use this zip compression when makin the pptx zip file; overwrite: Whether to silntly overwrite existing output file if it already exists;
verbose: Verbosity level; on_error: What to do when error encountered

returns the filename of the newly generated pptx file 
'''
def downsize_pptx_images(filename,fname_filter=None,fsize_filter=int(0.5*2**20),convert_to = "png",img_max_size = 2048,quality = 90,optimize = True,img_mode = None,
	fill_color = None,outputfn_fmt = '{fnroot}.downsized.pptx',compress_type=zipfile.ZIP_DEFLATED,wait_before_zip = False,overwrite = None,on_error = 'raise',verbose = 2):
		




'''
Here we are creating an argument parser object which is required to handle command line arguments

'''
def get_argparser(defaults=None):		#defaults is a dictionary containing default values of cmd line arguments
	if defaults is None:
		spec = inspect.getfullargspec(downsize_pptx_images)		#this fucntion is used to retrieve the full argument specs of the object specified
		defaults = dict(zip((spec.args + spec.kwonlyargs)[::-1], (spec.defaults + (spec.kwonlydefaults or ()))[::-1]))		#defaults is a dictionary made by combining the argument names and default values obtained from the above result
	ap = argparse.ArgumentParser(		#creating a command line argument parser
			description=(
					"Powepoint pptx downsizer"
					"Reduce the file size of Powepoint presentation by re-compressing images within the ppt file."
				),
			formatter_class = argparse.ArgumentDefaultsHelpFormatter
		)
	ap.add_argument('filename', help="Path to the PPT")
	ap.add_argument("--fname-filter", metavar="GLOB", default=defaults['fname_filter'], help=("Convert all images matching this filename pattern"))
	ap.add_argument("--fsize-filter",metavar="SIZE", default=defaults['fsize_filter'], help=("Convert all images with a current file size exceeding this limit"))
	ap.add_argument("--convert-to", metavar="IMAGE_FORMAT", default=defaults['convert_to'],help=("Convert images to this image fomat"))
	ap.add_argument("--img-max-size",metavar='PIXELS',default=defaults['img_max_size'], type=int, help=("If images are larger than this size (in px),reduce them and make it this size"))
	ap.add_argument("--img-mode", metavar="MODE", default=defaults['img_mode'], help=("Convert images to this image mode before saving them, e.g. 'RGB' - advanced option."))
    ap.add_argument("--fill-color", metavar="COLOR", default=defaults['fill_color'], help=("If converting image mode (e.g. from RGBA to RGB), use this color for transparent regions."))
    ap.add_argument("--quality", metavar="[1-100]", default=defaults['quality'], type=int, help=("Quality of converted images (only applies to jpeg output)."))
    ap.add_argument("--optimize", default=defaults['optimize'], action="store_true", dest="optimize", help=("Try to optimize the converted image output when saving.Optimizing the output may produce better images, but disabling it may make the conversion run faster. Enabled by default."))
    ap.add_argument("--no-optimize", default=not defaults['optimize'], action="store_false", dest="optimize", help=("Disable optimization."))
    
    # pptx output options:
    ap.add_argument("--outputfn_fmt", metavar="FORMAT-STRING", default=defaults['outputfn_fmt'], help=("How to format the downsized presentation pptx filename\n Slightly advanced, uses python string formatting."))
    ap.add_argument("--overwrite", default=defaults['overwrite'], action="store_true", help=("Whether to silently overwrite existing file if the output filename already exists."))
    ap.add_argument("--compress-type", metavar="ZIP-TYPE", default='ZIP_DEFLATED', help=("Which zip compression type to use, e.g. ZIP_DEFLATED, ZIP_BZIP2, or ZIP_LZMA."))
    ap.add_argument("--wait-before-zip", default=defaults['wait_before_zip'], action="store_true", help=("If this flag is specified, the program will wait after converting all images before re-zipping the output pptx file.You can use this to make manual changes to the presentation - advanced option."))
  
    # verbosity and other program/display behavior:
    ap.add_argument("--on-error", metavar="DO-WHAT", default=defaults['on_error'], help=("What to do if the program encounters any errors during execution.`continue` will cause the program to continue even if one or more images fails to be converted."))
    ap.add_argument("--verbose", metavar="[0-5]", default=defaults['verbose'], type=int, help=("Increase or decrease the 'verbosity' of the program,i.e. how much information it prints about the process."))
    # ap.add_argument("--open-pptx", default=defaults['open_pptx'], action="store_true")
    return ap
***


	return ap

'''
parse_args takes in the command line arguments checks if a size has been provided
if yes, then convert to int; else, ValueError. Check if a compression type has been provided 
if yes, then do the appropriate step
'''
def parse_args(argv=None, defaults=None):		#both the parameters are optional
	ap = get_argparser(defaults=defaults)
	argns = ap.parse_args(argv)

	if argns.fsize_filter:		#fsize_filter is file size given as command line parameter
		try:
			argns.fsize_filter = convert_str_to_int(argns.fsize_filter)
		except ValuError:
			ap.print_usage()		#prints the usage of the command
			print("error:fsize_filter must be numeric, is %r % argns.fsize_filter")
	if argns.compress_type and isinstance(argns.compress_type,str):		#argns.compress_type is the type of compression used ++++ isinstance() checks if arg1 is not empty and if it is a string or not
		argns.compres_type = getattr(zipfile,argns.compress_type)

	return argns

def cli(argv=None):		#default value is None
	argns = parse_args(argv)	#number of arguments
	params = vars(argns)		#converts argns into a dictionary
	if argns.verbose and argns.verbose > 2:
		print("parameters:")
		print(yaml.dump(params, default_flow_style = False))	#converts the dictionary into a YAML string represenation 
	downsize_pptx_images(**params)


if __name__ == '__main__':
	cli()