#include <stdio.h>
#include <stdlib.h>
#include <stdint.h>
#include <sys/stat.h>

#include "src/util/stream.h"
#include "src/lzcomp/liblzcomp.h"
#include "src/ctf/parseCTF.h"
#include "src/ctf/SFNTContainer.h"

static uint32_t read_u32le(const uint8_t *bytes)
{
  return (uint32_t)bytes[0] |
         ((uint32_t)bytes[1] << 8) |
         ((uint32_t)bytes[2] << 16) |
         ((uint32_t)bytes[3] << 24);
}

static uint16_t read_u16le(const uint8_t *bytes)
{
  return (uint16_t)(bytes[0] | (bytes[1] << 8));
}

static void print_eot_error(enum EOTError error)
{
  switch (error)
  {
  case EOT_SUCCESS:
    break;
  case EOT_INSUFFICIENT_BYTES:
    fprintf(stderr, "The embedded font file appears truncated.\n");
    break;
  case EOT_BOGUS_STRING_SIZE:
  case EOT_CORRUPT_FILE:
  case EOT_MTX_ERROR:
    fprintf(stderr, "The embedded font file appears corrupt or unsupported.\n");
    break;
  case EOT_CANT_ALLOCATE_MEMORY:
    fprintf(stderr, "Unable to allocate enough memory.\n");
    break;
  case EOT_FWRITE_ERROR:
  case EOT_OTHER_STDLIB_ERROR:
    fprintf(stderr, "A filesystem error occurred while writing the output font.\n");
    break;
  default:
    fprintf(stderr, "libeot returned error code %d.\n", error);
    break;
  }
}

static void usage(const char *program_name)
{
  fprintf(stderr, "Usage: %s input.fntdata output.ttf\n", program_name);
}

static int looks_like_sfnt(const uint8_t *buffer, unsigned size)
{
  if (size < 4)
  {
    return 0;
  }

  return
    (buffer[0] == 0x00 && buffer[1] == 0x01 && buffer[2] == 0x00 && buffer[3] == 0x00) ||
    (buffer[0] == 0x4F && buffer[1] == 0x54 && buffer[2] == 0x54 && buffer[3] == 0x4F) ||
    (buffer[0] == 0x74 && buffer[1] == 0x74 && buffer[2] == 0x63 && buffer[3] == 0x66) ||
    (buffer[0] == 0x74 && buffer[1] == 0x72 && buffer[2] == 0x75 && buffer[3] == 0x65);
}

int main(int argc, char **argv)
{
  if (argc != 3)
  {
    usage(argv[0]);
    return 1;
  }

  struct stat st;
  if (stat(argv[1], &st) != 0)
  {
    fprintf(stderr, "The file %s could not be opened.\n", argv[1]);
    return 1;
  }

  FILE *input_file = fopen(argv[1], "rb");
  if (input_file == NULL)
  {
    fprintf(stderr, "The file %s could not be opened.\n", argv[1]);
    return 1;
  }

  uint8_t *buffer = (uint8_t *)malloc((size_t)st.st_size);
  if (buffer == NULL)
  {
    fclose(input_file);
    fprintf(stderr, "Unable to allocate enough memory.\n");
    return 1;
  }

  size_t bytes_read = fread(buffer, 1, (size_t)st.st_size, input_file);
  fclose(input_file);
  if (bytes_read != (size_t)st.st_size)
  {
    free(buffer);
    fprintf(stderr, "The file %s could not be fully read.\n", argv[1]);
    return 1;
  }

  if (st.st_size < 36)
  {
    free(buffer);
    fprintf(stderr, "The file %s is too small to be a valid WPS EOT font.\n", argv[1]);
    return 1;
  }

  uint32_t total_size = read_u32le(buffer);
  uint32_t font_data_size = read_u32le(buffer + 4);
  uint32_t version = read_u32le(buffer + 8);
  uint16_t magic = read_u16le(buffer + 34);

  if (total_size != (uint32_t)st.st_size ||
      font_data_size == 0 ||
      total_size < font_data_size ||
      magic != 0x504C ||
      (version != 0x00010000 && version != 0x00020001 && version != 0x00020002))
  {
    free(buffer);
    fprintf(stderr, "The file %s is not a supported WPS / EOT embedded font.\n", argv[1]);
    return 1;
  }

  uint32_t mtx_offset = total_size - font_data_size;
  struct Stream mtx_stream = constructStream(buffer + mtx_offset, font_data_size);
  uint8_t *streams_out[3] = {NULL, NULL, NULL};
  unsigned streams_size[3] = {0, 0, 0};
  enum EOTError result = unpackMtx(&mtx_stream, font_data_size, streams_out, streams_size);
  free(buffer);

  if (result != EOT_SUCCESS)
  {
    print_eot_error(result);
    return 1;
  }

  uint8_t *font_out = NULL;
  unsigned font_out_size = 0;

  if (looks_like_sfnt(streams_out[0], streams_size[0]))
  {
    font_out = streams_out[0];
    font_out_size = streams_size[0];
    streams_out[0] = NULL;
  }
  else
  {
    struct Stream ctf_streams[3];
    struct Stream *ctf_stream_ptrs[3];
    struct SFNTContainer *container = NULL;
    unsigned i;

    for (i = 0; i < 3; ++i)
    {
      ctf_streams[i] = constructStream(streams_out[i], streams_size[i]);
      ctf_stream_ptrs[i] = &ctf_streams[i];
    }

    result = parseCTF(ctf_stream_ptrs, &container);
    if (result != EOT_SUCCESS)
    {
      for (i = 0; i < 3; ++i)
      {
        free(streams_out[i]);
      }
      print_eot_error(result);
      return 1;
    }

    result = dumpContainer(container, &font_out, &font_out_size);
    freeContainer(container);
    for (i = 0; i < 3; ++i)
    {
      free(streams_out[i]);
      streams_out[i] = NULL;
    }

    if (result != EOT_SUCCESS)
    {
      print_eot_error(result);
      return 1;
    }
  }

  FILE *output_file = fopen(argv[2], "wb");
  if (output_file == NULL)
  {
    free(font_out);
    fprintf(stderr, "The file %s could not be opened for writing.\n", argv[2]);
    return 1;
  }

  size_t bytes_written = fwrite(font_out, 1, (size_t)font_out_size, output_file);
  fclose(output_file);
  free(font_out);

  if (bytes_written != font_out_size)
  {
    fprintf(stderr, "The file %s could not be fully written.\n", argv[2]);
    return 1;
  }

  {
    unsigned i;
    for (i = 0; i < 3; ++i)
    {
      free(streams_out[i]);
    }
  }

  return 0;
}
