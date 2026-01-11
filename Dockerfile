FROM php:8.1-apache

# Install system dependencies
RUN apt-get update && apt-get install -y \
    git \
    curl \
    zip \
    unzip \
    libxml2-dev \
    libzip-dev \
    zlib1g-dev \
    && docker-php-ext-install -j$(nproc) \
    dom \
    zip \
    && rm -rf /var/lib/apt/lists/*

# Enable Apache mod_rewrite
RUN a2enmod rewrite

# Install Composer
COPY --from=composer:latest /usr/bin/composer /usr/bin/composer

# Set working directory
WORKDIR /var/www/html

# Copy application files
COPY . .

# Install PHP dependencies
RUN composer install --no-dev --optimize-autoloader

# Create output directory and set permissions
RUN mkdir -p /var/www/html/output \
    && chown -R www-data:www-data /var/www/html \
    && chmod -R 755 /var/www/html \
    && chmod -R 775 /var/www/html/output

# Expose port
EXPOSE 80

# Start Apache in foreground
CMD ["apache2-foreground"]
