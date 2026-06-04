-- Film Rental (Pagila) Schema: DVD rental store
-- Source: github.com/devrimgunduz/pagila
-- PostgreSQL dialect

CREATE TABLE customer (
    customer_id SERIAL NOT NULL,
    store_id integer NOT NULL REFERENCES store(store_id),
    first_name text NOT NULL,
    last_name text NOT NULL,
    email text,
    address_id integer NOT NULL REFERENCES address(address_id),
    activebool boolean DEFAULT true NOT NULL,
    create_date date DEFAULT CURRENT_DATE NOT NULL,
    last_update timestamp with time zone DEFAULT now(),
    active integer
);

CREATE TABLE actor (
    actor_id SERIAL NOT NULL,
    first_name text NOT NULL,
    last_name text NOT NULL,
    last_update timestamp with time zone DEFAULT now() NOT NULL
);

CREATE TABLE category (
    category_id SERIAL NOT NULL,
    name text NOT NULL,
    last_update timestamp with time zone DEFAULT now() NOT NULL
);

CREATE TABLE film (
    film_id SERIAL NOT NULL,
    title text NOT NULL,
    description text,
    release_year integer,
    language_id integer NOT NULL REFERENCES language(language_id),
    original_language_id integer REFERENCES language(language_id),
    rental_duration smallint DEFAULT 3 NOT NULL,
    rental_rate numeric(4,2) DEFAULT 4.99 NOT NULL,
    length smallint,
    replacement_cost numeric(5,2) DEFAULT 19.99 NOT NULL,
    rating VARCHAR(10) DEFAULT 'G',
    last_update timestamp with time zone DEFAULT now() NOT NULL,
    special_features text
);

CREATE TABLE film_actor (
    actor_id integer NOT NULL REFERENCES actor(actor_id),
    film_id integer NOT NULL REFERENCES film(film_id),
    last_update timestamp with time zone DEFAULT now() NOT NULL
);

CREATE TABLE film_category (
    film_id integer NOT NULL REFERENCES film(film_id),
    category_id integer NOT NULL REFERENCES category(category_id),
    last_update timestamp with time zone DEFAULT now() NOT NULL
);

CREATE TABLE address (
    address_id SERIAL NOT NULL,
    address text NOT NULL,
    address2 text,
    district text NOT NULL,
    city_id integer NOT NULL REFERENCES city(city_id),
    postal_code text,
    phone text NOT NULL,
    last_update timestamp with time zone DEFAULT now() NOT NULL
);

CREATE TABLE city (
    city_id SERIAL NOT NULL,
    city text NOT NULL,
    country_id integer NOT NULL REFERENCES country(country_id),
    last_update timestamp with time zone DEFAULT now() NOT NULL
);

CREATE TABLE country (
    country_id SERIAL NOT NULL,
    country text NOT NULL,
    last_update timestamp with time zone DEFAULT now() NOT NULL
);

CREATE TABLE inventory (
    inventory_id SERIAL NOT NULL,
    film_id integer NOT NULL REFERENCES film(film_id),
    store_id integer NOT NULL REFERENCES store(store_id),
    last_update timestamp with time zone DEFAULT now() NOT NULL
);

CREATE TABLE language (
    language_id SERIAL NOT NULL,
    name character(20) NOT NULL,
    last_update timestamp with time zone DEFAULT now() NOT NULL
);

CREATE TABLE payment (
    payment_id SERIAL NOT NULL,
    customer_id integer NOT NULL REFERENCES customer(customer_id),
    staff_id integer NOT NULL REFERENCES staff(staff_id),
    rental_id integer NOT NULL REFERENCES rental(rental_id),
    amount numeric(5,2) NOT NULL,
    payment_date timestamp with time zone NOT NULL
);

CREATE TABLE rental (
    rental_id SERIAL NOT NULL,
    rental_date timestamp with time zone NOT NULL,
    inventory_id integer NOT NULL REFERENCES inventory(inventory_id),
    customer_id integer NOT NULL REFERENCES customer(customer_id),
    return_date timestamp with time zone,
    staff_id integer NOT NULL REFERENCES staff(staff_id),
    last_update timestamp with time zone DEFAULT now() NOT NULL
);

CREATE TABLE staff (
    staff_id SERIAL NOT NULL,
    first_name text NOT NULL,
    last_name text NOT NULL,
    address_id integer NOT NULL REFERENCES address(address_id),
    email text,
    store_id integer NOT NULL REFERENCES store(store_id),
    active boolean DEFAULT true NOT NULL,
    username text NOT NULL,
    password text,
    last_update timestamp with time zone DEFAULT now() NOT NULL,
    picture bytea
);

CREATE TABLE store (
    store_id SERIAL NOT NULL,
    manager_staff_id integer NOT NULL,
    address_id integer NOT NULL REFERENCES address(address_id),
    last_update timestamp with time zone DEFAULT now() NOT NULL
);
