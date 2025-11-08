"use client";

import { useState, useEffect, useRef } from "react";
import { Button } from "@/components/ui/button";
import { Download, ChevronLeft, ChevronRight, FileText } from "lucide-react";

interface PptxViewerProps {
  fileUrl: string;
}

interface SlideInfo {
  slideNumber: number;
  title: string;
  text: string[];
}

export function PptxViewer({ fileUrl }: PptxViewerProps) {
  const [currentSlide, setCurrentSlide] = useState(0);
  const [slides, setSlides] = useState<SlideInfo[]>([]);
  const [slideImages, setSlideImages] = useState<Record<number, string>>({});
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [hasVisualSlides, setHasVisualSlides] = useState(false);
  const imagesRef = useRef<Record<number, string>>({});
  const hasVisualRef = useRef(false);

  // Keep refs in sync with state
  useEffect(() => {
    imagesRef.current = slideImages;
    hasVisualRef.current = hasVisualSlides;
  }, [slideImages, hasVisualSlides]);

  useEffect(() => {
    if (!fileUrl) return;

    let isMounted = true;
    const loadSlides = async () => {
      // Clear existing images when fileUrl changes (fresh load after edit)
      console.log(
        `🔄 fileUrl changed, clearing images and reloading. fileUrl: ${fileUrl}`,
      );
      if (isMounted) {
        setSlideImages({});
        setHasVisualSlides(false);
        imagesRef.current = {};
        hasVisualRef.current = false;
      }

      // Only show loading if we don't have slides yet
      if (slides.length === 0) {
        setLoading(true);
      }

      try {
        // First, fetch slides info (with cache busting)
        const slidesResponse = await fetch(
          `http://localhost:3001/api/slides?t=${Date.now()}`,
          {
            cache: "no-store",
          },
        );

        if (slidesResponse.ok) {
          const slidesData = await slidesResponse.json();
          if (isMounted) {
            setSlides(slidesData.slides || []);
            setError(null);
          }
        } else {
          if (isMounted) {
            setError("Failed to load slides");
            setLoading(false);
          }
          return;
        }

        // Then try to load slide images (may take time if LibreOffice is converting)
        const loadImages = async (retries = 10) => {
          for (let i = 0; i < retries; i++) {
            if (!isMounted) return;

            try {
              const imagesResponse = await fetch(
                `http://localhost:3001/api/slide-images-all?t=${Date.now()}`,
                { cache: "no-store" },
              );

              if (imagesResponse.ok) {
                const imagesData = await imagesResponse.json();
                console.log(`Images API response:`, {
                  hasSlides: !!imagesData.slides,
                  count: imagesData.count,
                  hasImages: imagesData.hasImages,
                  slideKeys: imagesData.slides
                    ? Object.keys(imagesData.slides)
                    : [],
                });

                if (
                  imagesData.slides &&
                  Object.keys(imagesData.slides).length > 0
                ) {
                  // Only update if we're still mounted
                  if (isMounted) {
                    // Replace images (don't merge) to ensure we get fresh images after edits
                    console.log(
                      `📸 Received ${imagesData.count} images from backend:`,
                      Object.keys(imagesData.slides),
                    );
                    // Log first few characters of base64 to verify it's different
                    if (imagesData.slides[2]) {
                      const preview = imagesData.slides[2].substring(0, 50);
                      console.log(
                        `Slide 2 image preview (first 50 chars): ${preview}...`,
                      );
                    }
                    setSlideImages(imagesData.slides);
                    imagesRef.current = imagesData.slides;
                    setHasVisualSlides(true);
                    hasVisualRef.current = true;
                    console.log(
                      `✅ Loaded ${imagesData.count} fresh visual slides (attempt ${i + 1})`,
                    );
                  }
                  return; // Success, stop retrying
                } else if (i === retries - 1) {
                  // If we have existing images in ref, keep them (don't clear)
                  if (isMounted) {
                    const currentImages = imagesRef.current;
                    if (Object.keys(currentImages).length > 0) {
                      console.log(
                        `Keeping ${Object.keys(currentImages).length} existing images (no new ones found)`,
                      );
                      // Don't update state - keep existing images
                    } else {
                      console.warn(
                        `No slide images found after ${retries} attempts`,
                      );
                    }
                  }
                }
              } else {
                console.error(
                  `Images API error: ${imagesResponse.status} ${imagesResponse.statusText}`,
                );
              }

              // If no images yet and we have retries left, wait and try again
              if (i < retries - 1) {
                await new Promise((resolve) => setTimeout(resolve, 2000));
              }
            } catch (err) {
              console.error(`Error loading images (attempt ${i + 1}):`, err);
              if (i < retries - 1) {
                await new Promise((resolve) => setTimeout(resolve, 2000));
              }
            }
          }
        };

        // Load images in the background (non-blocking)
        // Don't clear existing images while loading
        loadImages().catch((err) => {
          console.error("Failed to load slide images after retries:", err);
        });
      } catch (err) {
        console.error("Error loading slides:", err);
        if (isMounted) {
          setError("Failed to load slides");
        }
      } finally {
        if (isMounted) {
          setLoading(false);
        }
      }
    };

    loadSlides();

    // Cleanup function to prevent state updates after unmount
    return () => {
      isMounted = false;
    };
  }, [fileUrl]); // Only depend on fileUrl, not on slides or slideImages

  const handleDownload = () => {
    const link = document.createElement("a");
    link.href = fileUrl;
    link.download = "presentation.pptx";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const handlePrevious = () => {
    setCurrentSlide((prev) => Math.max(0, prev - 1));
  };

  const handleNext = () => {
    setCurrentSlide((prev) => Math.min(slides.length - 1, prev + 1));
  };

  if (loading) {
    return (
      <div className="flex h-full items-center justify-center">
        <div className="text-center">
          <p className="mb-2">Loading presentation...</p>
          <p className="text-sm text-muted-foreground">Please wait</p>
        </div>
      </div>
    );
  }

  if (error || slides.length === 0) {
    return (
      <div className="flex h-full flex-col bg-background p-6">
        <div className="mb-4 flex items-center justify-between">
          <div>
            <h2 className="text-lg font-semibold">Presentation Viewer</h2>
            <p className="text-sm text-muted-foreground">
              {error || "No slides found"}
            </p>
          </div>
          <Button onClick={handleDownload} variant="outline" size="sm">
            <Download className="mr-2 h-4 w-4" />
            Download
          </Button>
        </div>
        <div className="flex flex-1 items-center justify-center rounded-lg border border-dashed p-8">
          <div className="text-center">
            <FileText className="mx-auto mb-4 h-16 w-16 text-muted-foreground" />
            <p className="mb-4 text-sm text-muted-foreground">
              {error ||
                "Unable to load slides. You can download the file to view it."}
            </p>
            <Button onClick={handleDownload} variant="outline">
              <Download className="mr-2 h-4 w-4" />
              Download Presentation
            </Button>
          </div>
        </div>
      </div>
    );
  }

  const currentSlideData = slides[currentSlide];

  return (
    <div className="flex h-full flex-col bg-background">
      {/* Header */}
      <div className="flex items-center justify-between border-b bg-background p-4">
        <div className="flex items-center gap-4">
          <div>
            <h2 className="text-lg font-semibold">Presentation Viewer</h2>
            <p className="text-sm text-muted-foreground">
              Slide {currentSlide + 1} of {slides.length}
            </p>
          </div>
        </div>
        <div className="flex gap-2">
          <Button
            onClick={handlePrevious}
            disabled={currentSlide === 0}
            variant="outline"
            size="sm"
          >
            <ChevronLeft className="h-4 w-4" />
          </Button>
          <Button
            onClick={handleNext}
            disabled={currentSlide === slides.length - 1}
            variant="outline"
            size="sm"
          >
            <ChevronRight className="h-4 w-4" />
          </Button>
          <Button onClick={handleDownload} variant="default" size="sm">
            <Download className="mr-2 h-4 w-4" />
            Download
          </Button>
        </div>
      </div>

      {/* Content */}
      <div className="flex-1 overflow-auto p-6">
        <div className="mx-auto max-w-6xl">
          {/* Visual slide display */}
          <div className="mb-6 flex items-center justify-center rounded-lg border-2 bg-card p-8 shadow-lg">
            {hasVisualSlides && slideImages[currentSlide + 1] ? (
              <div className="flex w-full flex-col items-center">
                <img
                  key={`slide-${currentSlide}-${fileUrl}`}
                  src={slideImages[currentSlide + 1]}
                  alt={`Slide ${currentSlide + 1}`}
                  className="max-w-full rounded border bg-white shadow-md"
                  style={{
                    maxHeight: "70vh",
                    objectFit: "contain",
                  }}
                  onError={(e) => {
                    console.error(
                      `Error loading image for slide ${currentSlide + 1}`,
                    );
                    // Don't disable all visual slides, just log the error
                    // The image will just not display, but other slides can still show
                  }}
                />
                <div className="mt-4 text-center">
                  <p className="text-sm text-muted-foreground">
                    Slide {currentSlide + 1} of {slides.length}
                  </p>
                </div>
              </div>
            ) : (
              <div className="w-full max-w-4xl">
                {/* Fallback to text-only display */}
                <div className="mb-6 border-b pb-4">
                  <div className="mb-2 flex items-center justify-between">
                    <span className="text-sm font-medium text-muted-foreground">
                      Slide {currentSlide + 1} of {slides.length}
                    </span>
                    <span className="rounded-full bg-primary/10 px-3 py-1 text-xs font-medium text-primary">
                      {currentSlideData.slideNumber}
                    </span>
                  </div>
                  <h2 className="text-3xl leading-tight font-bold">
                    {currentSlideData.title}
                  </h2>
                </div>
                <div className="space-y-4">
                  {currentSlideData.text.slice(1).length > 0 ? (
                    currentSlideData.text.slice(1).map((text, index) => (
                      <div
                        key={index}
                        className="rounded-lg bg-background/50 p-4 text-lg leading-relaxed"
                      >
                        {text}
                      </div>
                    ))
                  ) : (
                    <p className="text-lg text-muted-foreground italic">
                      No additional content on this slide
                    </p>
                  )}
                </div>
                {!hasVisualSlides && (
                  <div className="mt-4 rounded-lg border border-yellow-200 bg-yellow-50 p-4">
                    <p className="text-sm text-yellow-800">
                      <strong>Note:</strong> Visual slide rendering requires
                      LibreOffice to be installed. Install it to see slides with
                      images and colors. For now, showing text content only.
                    </p>
                  </div>
                )}
              </div>
            )}
          </div>

          {/* Slide thumbnails */}
          <div className="mt-8">
            <div className="mb-4 flex items-center justify-between">
              <p className="text-sm font-semibold text-foreground">
                All Slides ({slides.length})
              </p>
              <Button
                onClick={handleDownload}
                variant="ghost"
                size="sm"
                className="text-xs"
              >
                <Download className="mr-2 h-3 w-3" />
                Download All
              </Button>
            </div>
            <div className="grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-4">
              {slides.map((slide, index) => {
                const hasThumbnail = hasVisualSlides && slideImages[index + 1];
                return (
                  <button
                    key={slide.slideNumber}
                    onClick={() => setCurrentSlide(index)}
                    className={`group relative overflow-hidden rounded-lg border-2 transition-all ${
                      index === currentSlide
                        ? "border-primary bg-primary/10 shadow-md ring-2 ring-primary/20"
                        : "border-border bg-card hover:border-primary/50 hover:bg-accent hover:shadow-sm"
                    }`}
                  >
                    {hasThumbnail && slideImages[index + 1] ? (
                      <div className="relative">
                        <img
                          key={`thumb-${index}-${slide.slideNumber}`}
                          src={slideImages[index + 1]}
                          alt={`Slide ${slide.slideNumber} thumbnail`}
                          className="h-32 w-full object-cover"
                          onError={(e) => {
                            console.error(
                              `Error loading thumbnail for slide ${index + 1}`,
                            );
                          }}
                        />
                        <div className="absolute right-0 bottom-0 left-0 bg-black/60 px-2 py-1">
                          <p className="text-xs font-semibold text-white">
                            Slide {slide.slideNumber}
                          </p>
                        </div>
                        {index === currentSlide && (
                          <div className="absolute top-2 right-2 h-3 w-3 rounded-full bg-primary ring-2 ring-white" />
                        )}
                      </div>
                    ) : (
                      <div className="p-4 text-left">
                        <div className="mb-2 flex items-center justify-between">
                          <span className="text-xs font-medium text-muted-foreground">
                            #{slide.slideNumber}
                          </span>
                          {index === currentSlide && (
                            <div className="h-2 w-2 rounded-full bg-primary" />
                          )}
                        </div>
                        <p className="line-clamp-3 text-sm leading-snug font-semibold">
                          {slide.title}
                        </p>
                        {slide.text.length > 1 && (
                          <p className="mt-2 line-clamp-2 text-xs text-muted-foreground">
                            {slide.text.slice(1).join(" ")}
                          </p>
                        )}
                      </div>
                    )}
                  </button>
                );
              })}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
