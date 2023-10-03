def placeholder_identifier(slide):
    for shape in slide.shapes:
        if shape.is_placeholder:
            phf = shape.placeholder_format
        print('%d, %s' % (phf.idx, phf.type))